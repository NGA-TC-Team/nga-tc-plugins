"""
inject_data.py
프로파일 JSON + HTML 템플릿을 받아 최종 대시보드 HTML을 생성한다.

사용법:
  python3 inject_data.py <profile.json> <template.html> <output.html> \
      --title "..." --source "원본.csv" --gen-date "2026-04-22"

KPI·차트·테이블 설정은 profile["suggest"]를 기본으로 쓰되, 사용자 커스터마이징을 허용한다.
사용자가 승인한 최종 suggest 구조를 profile.json에 반영한 뒤 이 스크립트를 실행한다.
"""
import sys, json, os, statistics, argparse
from collections import Counter, defaultdict

def build_kpi_config(profile):
    """KPI 카드 설정을 최종 텍스트 값으로 변환"""
    out = []
    total = profile["total"]
    for k in profile["suggest"]["kpis"]:
        label = k["label"]
        value = k.get("value")
        unit = k.get("unit","")
        accent = k.get("accent","primary")
        hint = ""
        if k["kind"] == "count":
            out.append({"label":label,"value":str(value),"unit":unit,"hint":"전체 유효 응답","accent":accent}); continue
        if k["kind"] == "avg" and k.get("col"):
            col = k["col"]; s = profile["columns"][col]["stats"]
            out.append({"label":label,"value":f'{s.get("avg"):.2f}' if s.get("avg") is not None else '-',
                        "unit":"","hint":f'최대 {s.get("max")} · 최소 {s.get("min")}',"accent":accent}); continue
        if k["kind"] == "max" and k.get("col"):
            col = k["col"]; s = profile["columns"][col]["stats"]
            out.append({"label":label,"value":str(s.get("max","-")),"unit":"","hint":f'합계 {s.get("sum","-")}',"accent":accent}); continue
        if k["kind"] == "cat_count" and k.get("col"):
            col = k["col"]; m = profile["columns"][col]
            cnt = m["stats"].get("counts", {})
            top = " · ".join(list(cnt.keys())[:3])
            out.append({"label":label,"value":str(m["unique_count"]),"unit":"개","hint":top,"accent":accent}); continue
        if k["kind"] == "ratio":
            # {col, cmp: '>=4'} 만족률 등
            col = k["col"]; thr = k["threshold"]
            vals = [float(r.get(col)) for r in profile["rows"] if r.get(col) not in (None,"")]
            if vals:
                rate = sum(1 for v in vals if v >= thr) / len(vals) * 100
                out.append({"label":label,"value":f'{rate:.1f}',"unit":"%","hint":f"{thr}점 이상 기준","accent":accent})
            continue
        if k["kind"] == "count_filter":
            # text 컬럼에 값이 있는 행 수
            col = k["col"]
            cnt = sum(1 for r in profile["rows"] if r.get(col) and str(r[col]).strip() and str(r[col]).strip() not in (k.get("exclude") or []))
            out.append({"label":label,"value":str(cnt),"unit":"건","hint":k.get("hint",""),"accent":accent})
            continue
        # fallback
        out.append({"label":label,"value":str(value or '-'),"unit":unit,"hint":"","accent":accent})
    return out

def build_chart_config(profile):
    rows = profile["rows"]
    out = {"top": [], "full": None}
    for ch in profile["suggest"]["charts"]["top"]:
        out["top"].append(_make_chart(ch, rows, profile["columns"]))
    full = profile["suggest"]["charts"]["full"]
    if full:
        out["full"] = _make_chart(full, rows, profile["columns"])
    return out

def _num(v):
    if v is None: return None
    try: return float(str(v).replace(",",""))
    except: return None

def _make_chart(ch, rows, cols_meta):
    kind = ch["kind"]
    if kind == "distribution":
        col = ch["col"]
        vals = [_num(r.get(col)) for r in rows]
        vals = [int(round(v)) for v in vals if v is not None]
        c = Counter(vals)
        keys = sorted(c.keys())
        colors = _score_colors(keys)
        return {
            "type":"bar","title":ch["title"],"subtitle":"값별 응답 수",
            "data":{"labels":[str(k) for k in keys],
                    "datasets":[{"label":"응답 수","data":[c[k] for k in keys],"backgroundColor":colors,"borderRadius":8,"borderSkipped":False}]},
            "showLegend": False,
        }
    if kind == "group_avg":
        g = ch["group"]; v = ch["value"]
        buckets = defaultdict(list)
        for r in rows:
            gv = r.get(g); vv = _num(r.get(v))
            if gv is None or vv is None: continue
            buckets[str(gv).strip()].append(vv)
        stats = sorted([(k, len(lst), sum(lst)/len(lst)) for k,lst in buckets.items()], key=lambda x: -x[1])
        labels = [s[0] for s in stats]
        return {
            "type":"bar","title":ch["title"],"subtitle":"응답 수 및 평균","indexAxis":"y",
            "data":{"labels":labels,"datasets":[
                {"label":"응답 수","data":[s[1] for s in stats],"backgroundColor":"#2563eb","borderRadius":6,"yAxisID":"y"},
                {"label":"평균","data":[round(s[2],2) for s in stats],"backgroundColor":"#fbbf24","borderRadius":6,"yAxisID":"y1"}
            ]},
            "scales":{
                "x":{"grid":{"color":"#f1f5f9"},"ticks":{"font":{"family":"Pretendard Variable","size":11}},"beginAtZero":True},
                "y":{"grid":{"display":False},"ticks":{"font":{"family":"Pretendard Variable","size":12,"weight":600}}},
                "y1":{"display":False}
            }
        }
    if kind == "trend":
        x = ch["x"]; y = ch["y"]
        buckets = defaultdict(list)
        for r in rows:
            xv = r.get(x); yv = _num(r.get(y))
            if xv is None or yv is None: continue
            buckets[str(xv).strip()].append(yv)
        keys = sorted(buckets.keys())
        counts = [len(buckets[k]) for k in keys]
        avgs = [round(sum(buckets[k])/len(buckets[k]), 2) for k in keys]
        ymin = min(avgs) if avgs else 0
        ymax = max(avgs) if avgs else 5
        return {
            "type":"bar","title":ch["title"],"subtitle":"기간별 응답 수 및 평균",
            "data":{"labels":keys,"datasets":[
                {"type":"bar","label":"응답 수","data":counts,"backgroundColor":"#e0e7ff","borderColor":"#818cf8","borderWidth":1,"borderRadius":4,"yAxisID":"y"},
                {"type":"line","label":"평균","data":avgs,"borderColor":"#2563eb","backgroundColor":"rgba(37,99,235,.1)","borderWidth":2,"pointRadius":4,"pointBackgroundColor":"#2563eb","tension":0.35,"yAxisID":"y1","fill":True}
            ]},
            "scales":{
                "x":{"grid":{"display":False},"ticks":{"font":{"family":"Pretendard Variable","size":11}}},
                "y":{"position":"left","grid":{"color":"#f1f5f9"},"ticks":{"font":{"family":"Pretendard Variable","size":11},"precision":0},"beginAtZero":True,"title":{"display":True,"text":"응답 수","font":{"family":"Pretendard Variable","size":11}}},
                "y1":{"position":"right","grid":{"display":False},"ticks":{"font":{"family":"Pretendard Variable","size":11}},"min":max(0, ymin-0.5),"max":min(5, ymax+0.5) if ymax<=5 else ymax,"title":{"display":True,"text":"평균","font":{"family":"Pretendard Variable","size":11}}}
            }
        }
    return None

def _score_colors(keys):
    # 1~5 스케일이면 빨강→초록, 그 외는 단일
    palette5 = {1:"#ef4444",2:"#f97316",3:"#f59e0b",4:"#10b981",5:"#059669"}
    if all(k in palette5 for k in keys):
        return [palette5[k] for k in keys]
    return ["#2563eb"] * len(keys)

def build_column_meta(profile):
    meta = {}
    for col, m in profile["columns"].items():
        entry = {"type": m["type"], "label": col}
        # numeric 1~5 스케일은 score 렌더러 사용
        if m["type"] == "numeric":
            s = m["stats"]
            if s.get("min",0) >= 1 and s.get("max",0) <= 5:
                entry["render"] = "score"
        if m["type"] == "categorical":
            entry["render"] = "category"
        meta[col] = entry
    return meta

def render(profile_path, template_path, out_path, title, source, gen_date):
    with open(profile_path, 'r', encoding='utf-8') as f:
        profile = json.load(f)
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    sg = profile["suggest"]
    rows = profile["rows"]
    column_meta = build_column_meta(profile)
    kpi_config  = build_kpi_config(profile)
    chart_config= build_chart_config(profile)
    fb_cards = [{"col":f["col"],"title":f["title"],"emoji":f.get("emoji","💬"),
                 "cls":f.get("cls","info"),"minLen":f.get("minLen",5),
                 "metaCols":f.get("metaCols", sg["filter_cols"][:2])} for f in sg["feedback"]]

    replacements = {
        "__TITLE__": title,
        "__SOURCE_NAME__": source,
        "__GEN_DATE__": gen_date,
        "__DATA_PLACEHOLDER__": json.dumps(rows, ensure_ascii=False),
        "__COLUMN_META__": json.dumps(column_meta, ensure_ascii=False),
        "__KPI_CONFIG__": json.dumps(kpi_config, ensure_ascii=False),
        "__CHART_CONFIG__": json.dumps(chart_config, ensure_ascii=False),
        "__FILTER_COLS__": json.dumps(sg["filter_cols"], ensure_ascii=False),
        "__SEARCH_COLS__": json.dumps(sg["search_cols"], ensure_ascii=False),
        "__TABLE_COLS__": json.dumps(sg["table_cols"], ensure_ascii=False),
        "__FEEDBACK_CARDS__": json.dumps(fb_cards, ensure_ascii=False),
    }
    for k,v in replacements.items():
        html = html.replace(k, v)

    # 덮어쓰기 방지: 파일 있으면 _vN
    final = out_path
    n = 2
    while os.path.exists(final):
        base, ext = os.path.splitext(out_path)
        final = f"{base}_v{n}{ext}"
        n += 1

    with open(final, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] 대시보드 저장: {final}")
    print(f"  파일 크기: {os.path.getsize(final):,} bytes")
    # 플레이스홀더 잔존 체크
    leftover = [k for k in replacements.keys() if k in html]
    if leftover:
        print(f"  ⚠ 잔존 플레이스홀더: {leftover}")
    return final

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("profile")
    ap.add_argument("template")
    ap.add_argument("output")
    ap.add_argument("--title", default="Data Dashboard")
    ap.add_argument("--source", default="")
    ap.add_argument("--gen-date", default="")
    args = ap.parse_args()
    render(args.profile, args.template, args.output, args.title, args.source, args.gen_date)

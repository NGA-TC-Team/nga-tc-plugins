"""
profile_data.py
CSV 또는 XLSX 파일을 로드해 컬럼별 타입을 자동 분류한다.

사용법:
  python3 profile_data.py <path> [--out profile.json]

출력: 컬럼별 {name, label, type, unique_count, sample, stats} 구조의 JSON

타입 분류 규칙:
  numeric     : 95% 이상 숫자 파싱 가능
  date        : 80% 이상 ISO 날짜 또는 타임스탬프 패턴
  categorical : 고유값 수 <= 30 AND 고유값/전체 <= 0.3, 평균 글자수 < 30
  text        : 그 외 (평균 글자수 >= 20 자유텍스트)
  identifier  : unique == total (주로 ID, 이메일)
"""
import sys, re, json, csv, statistics
from collections import Counter

DATE_PAT = re.compile(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}')

def _num(v):
    if v is None: return None
    s = str(v).strip().replace(',','')
    if s == '' or s.lower() in ('n/a','na','null','none'): return None
    try: return float(s)
    except ValueError: return None

def _is_date(v):
    if v is None: return False
    s = str(v).strip()
    return bool(DATE_PAT.match(s))

def load_rows(path):
    """CSV 또는 XLSX를 dict list로 로드"""
    if path.lower().endswith('.csv'):
        with open(path, 'r', encoding='utf-8-sig', newline='') as f:
            # 구분자 자동 감지 시도
            sample = f.read(4096); f.seek(0)
            try: dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
            except: dialect = csv.excel
            reader = csv.DictReader(f, dialect=dialect)
            return [dict(r) for r in reader if any((v or '').strip() for v in r.values())]
    elif path.lower().endswith(('.xlsx','.xlsm')):
        try:
            import openpyxl
        except ImportError:
            raise RuntimeError("openpyxl 미설치: pip install openpyxl --break-system-packages")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        headers = [str(h) if h is not None else f'col_{i}' for i,h in enumerate(next(it))]
        rows = []
        for r in it:
            if all(v is None for v in r): continue
            rows.append({headers[i]: r[i] if i < len(r) else None for i in range(len(headers))})
        return rows
    else:
        raise ValueError(f"지원하지 않는 확장자: {path}")

def profile(rows):
    if not rows:
        return {"total": 0, "columns": {}}
    cols = list(rows[0].keys())
    profile = {"total": len(rows), "columns": {}}
    for col in cols:
        vals = [r.get(col) for r in rows]
        non_null = [v for v in vals if v is not None and str(v).strip() != '']
        n = len(non_null)
        if n == 0:
            profile["columns"][col] = {"type":"empty","label":col,"unique_count":0,"sample":[],"stats":{}}
            continue

        # 숫자 비율
        num_vals = [_num(v) for v in non_null]
        num_count = sum(1 for x in num_vals if x is not None)
        num_ratio = num_count / n

        # 날짜 비율
        date_count = sum(1 for v in non_null if _is_date(v))
        date_ratio = date_count / n

        # 고유값
        str_vals = [str(v).strip() for v in non_null]
        uniques = list(dict.fromkeys(str_vals))  # 순서 유지 dedup
        u = len(uniques)
        avg_len = statistics.mean(len(s) for s in str_vals)

        stats = {}
        # 우선순위: numeric > date > text(자유) > identifier(짧은 고유) > categorical
        if num_ratio >= 0.95:
            clean = [x for x in num_vals if x is not None]
            stats = {
                "min": min(clean), "max": max(clean),
                "avg": round(statistics.mean(clean), 3),
                "sum": round(sum(clean), 3),
                "distinct": len(set(clean)),
            }
            ctype = "numeric"
        elif date_ratio >= 0.8:
            ctype = "date"
            stats = {"min": min(str_vals), "max": max(str_vals)}
        elif avg_len >= 20:
            # 자유 텍스트 (고유도와 무관)
            ctype = "text"
            stats = {"avg_len": round(avg_len,1), "non_empty": n}
        elif u / n > 0.9 and avg_len < 20:
            # 짧고 거의 고유한 값 → ID (이메일, 사번 등)
            ctype = "identifier"
        elif u <= 50 and u / n <= 0.5:
            ctype = "categorical"
            stats = {"counts": dict(Counter(str_vals).most_common(20))}
        else:
            ctype = "text"
            stats = {"avg_len": round(avg_len,1), "non_empty": n}

        profile["columns"][col] = {
            "type": ctype,
            "label": col,
            "unique_count": u,
            "non_null": n,
            "sample": uniques[:5],
            "stats": stats,
        }
    return profile

def suggest(profile_data):
    """프로파일 기반 KPI·차트 추천"""
    cols = profile_data["columns"]
    numeric_cols   = [c for c,m in cols.items() if m["type"]=="numeric"]
    cat_cols       = [c for c,m in cols.items() if m["type"]=="categorical"]
    date_cols      = [c for c,m in cols.items() if m["type"]=="date"]
    text_cols      = [c for c,m in cols.items() if m["type"]=="text"]

    kpis = [{"label":"총 응답 수","kind":"count","value":profile_data["total"],"unit":"건","accent":"primary"}]
    for nc in numeric_cols[:2]:
        s = cols[nc]["stats"]
        kpis.append({"label":f"{nc} 평균","kind":"avg","col":nc,"value":s.get("avg"),"accent":"ok"})
        if s.get("max") is not None:
            kpis.append({"label":f"{nc} 최고","kind":"max","col":nc,"value":s.get("max"),"accent":"accent"})
    for cc in cat_cols[:1]:
        kpis.append({"label":f"{cc} 수","kind":"cat_count","col":cc,"value":cols[cc]["unique_count"],"accent":"purple"})

    charts = {"top":[], "full": None}
    # 분포 차트: 첫 번째 numeric
    if numeric_cols:
        charts["top"].append({"kind":"distribution","col":numeric_cols[0],"title":f"{numeric_cols[0]} 분포"})
    # 그룹별 평균: 첫 cat × 첫 numeric
    if cat_cols and numeric_cols:
        charts["top"].append({"kind":"group_avg","group":cat_cols[0],"value":numeric_cols[0],"title":f"{cat_cols[0]}별 {numeric_cols[0]} 평균"})
    # 추이: date × numeric
    if date_cols and numeric_cols:
        charts["full"] = {"kind":"trend","x":date_cols[0],"y":numeric_cols[0],"title":f"{date_cols[0]}별 추이"}

    feedback = [{"col":c,"title":c,"emoji":"💬","cls":"info","minLen":5} for c in text_cols[:3]]

    return {
        "kpis": kpis,
        "charts": charts,
        "feedback": feedback,
        "table_cols": list(cols.keys()),
        "filter_cols": cat_cols,
        "search_cols": text_cols,
    }

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: profile_data.py <path> [--out profile.json]", file=sys.stderr); sys.exit(1)
    path = sys.argv[1]
    out = "profile.json"
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out")+1]
    rows = load_rows(path)
    p = profile(rows)
    p["suggest"] = suggest(p)
    p["rows"] = rows  # 주입용 원시 데이터
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(p, f, ensure_ascii=False, indent=2, default=str)
    print(f"[OK] 프로파일 저장: {out}")
    print(f"  총 {p['total']}행 / {len(p['columns'])}컬럼")
    for col, meta in p["columns"].items():
        print(f"  - {col}: {meta['type']} (unique={meta['unique_count']})")

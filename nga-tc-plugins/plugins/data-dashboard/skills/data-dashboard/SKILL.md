---
name: data-dashboard
description: "CSV 또는 XLSX 파일을 업로드 받아 인터랙티브 HTML 대시보드(KPI 카드 + Chart.js 시각화 + 필터/정렬/검색 가능한 테이블)를 자동 생성하는 스킬. 설문 응답, 매출 데이터, HR/운영 데이터 등 tabular data를 받아 단일 파일 HTML로 변환한다. 'HTML 대시보드', '대시보드 만들어줘', '데이터 시각화', '설문 대시보드', 'csv 분석 대시보드', 'xlsx 대시보드', '인터랙티브 리포트', 'data dashboard', '데이터 대시보드', '결과 시각화' 요청 시 사용. 원본 파일을 절대 수정/삭제하지 않고 output/ 폴더에 [원본파일명]_dashboard.html로 저장한다."
---

# Data Dashboard Generator Skill

## 역할

너는 데이터 시각화 전문가다. 사용자가 업로드한 CSV 또는 XLSX 파일을 받아, 팀·경영진이 바로 볼 수 있는 수준의 깔끔하고 전문적인 인터랙티브 HTML 대시보드를 생성한다.

## 절대 규칙

1. **원본 파일 보존**: CSV/XLSX 원본을 절대 수정/삭제/이동하지 않는다. 읽기 전용으로만 접근한다.
2. **계획 먼저**: 데이터를 분석한 후 반드시 KPI·차트·섹션 구성 계획을 먼저 사용자에게 제시하고 **승인을 받은 뒤** HTML을 생성한다.
3. **단일 파일 HTML**: 모든 CSS/JS를 인라인으로, Chart.js·Pretendard는 CDN으로 가져와 단 하나의 .html 파일로 완성한다.
4. **덮어쓰기 금지**: 출력 파일명이 이미 존재하면 `_v2`, `_v3` 접미사를 붙인다.
5. **경로**: 결과물은 반드시 `output/` 폴더에 저장한다.

## 워크플로

### Step 1. 데이터 프로파일링

사용자가 CSV/XLSX를 제공하면, 먼저 Python으로 로드해 다음을 파악한다.

```python
# CSV
import csv
with open(path, encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

# XLSX
import openpyxl
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
rows = [dict(zip([c.value for c in ws[1]], [c.value for c in r])) for r in ws.iter_rows(min_row=2)]
```

각 컬럼을 4가지 타입으로 자동 분류한다:
- **numeric**: 숫자 (점수, 금액, 수량) → KPI 평균/합계·히스토그램·평균 비교
- **categorical**: 범주형 (부서, 성별, 지역) → 카운트·그룹별 집계
- **date**: 날짜/타임스탬프 → 시계열 추이
- **text**: 자유 텍스트 (피드백, 의견) → 피드백 카드 리스트

**판별 규칙**
- 모든 값이 숫자 파싱 가능 → numeric
- 고유값이 전체의 20% 미만 + 짧은 문자열 → categorical
- ISO 날짜 패턴 / 타임스탬프 → date
- 평균 글자수 20자 이상 → text

### Step 2. 계획 제시 (필수 승인 단계)

사용자에게 다음 형태로 계획을 보여주고 승인을 받는다:

```
📊 데이터 개요
- 총 N개 행, M개 컬럼
- 컬럼 타입: numeric(X), categorical(Y), date(Z), text(W)

📌 제안 KPI 카드 (4~5개)
1. 총 응답 수: N건
2. [numeric 컬럼] 평균: xx.x
3. [numeric 컬럼] 최대/최소
4. 범주 수: P개
5. 텍스트 의견: Q건

📈 제안 차트
1. [numeric 컬럼] 분포 (막대)
2. [categorical] 그룹별 [numeric] 평균 (가로 바)
3. [date]별 추이 (라인)

🗂 테이블 + 필터
- 필터: [categorical 컬럼들]
- 검색: [text 컬럼들]
- 정렬: 전체 컬럼

이대로 진행할까요? 조정하실 항목 있으면 알려주세요.
```

사용자가 "승인", "진행", "ok" 등으로 답하면 Step 3로. 수정 요청이 있으면 반영 후 재제시.

### Step 3. HTML 생성

`templates/dashboard_template.html`을 복사해 사용한다. 템플릿에는 다음 플레이스홀더가 있다:

- `__TITLE__`: 대시보드 제목 (원본 파일명 기반)
- `__SOURCE_NAME__`: 원본 파일명
- `__GEN_DATE__`: 생성일 (YYYY-MM-DD)
- `__DATA_PLACEHOLDER__`: JSON 데이터 배열
- `__COLUMN_META__`: 컬럼 메타 정보 JSON (타입·라벨)
- `__KPI_CONFIG__`: KPI 카드 설정 JSON
- `__CHART_CONFIG__`: 차트 설정 JSON
- `__FILTER_COLS__`: 필터 대상 컬럼 배열
- `__SEARCH_COLS__`: 검색 대상 컬럼 배열
- `__TABLE_COLS__`: 테이블 표시 컬럼 배열

템플릿 JavaScript가 위 설정을 읽어 자동으로 KPI·차트·테이블을 렌더링한다.

### Step 4. 저장

```
output/[원본파일명 확장자 제외]_dashboard.html
```

파일 존재 시:
```
output/[원본파일명]_dashboard_v2.html
output/[원본파일명]_dashboard_v3.html
...
```

### Step 5. 결과 전달

사용자에게 다음 형태로 전달:
1. `computer://` 링크로 HTML 즉시 열기 버튼
2. 핵심 인사이트 3~5줄 (평균/최대/특이값 등)
3. 대시보드 구성 요약 (KPI, 차트, 테이블 기능)

## 디자인 시스템 (표준)

### 컬러 팔레트

```css
--bg: #f5f7fb;           /* 배경 */
--surface: #ffffff;       /* 카드 */
--border: #e5e8ef;
--text: #0f172a;         /* 본문 */
--text-sub: #64748b;     /* 보조 */
--text-mute: #94a3b8;    /* 약한 */
--primary: #2563eb;      /* 포인트 */
--primary-soft: #eff4ff;
--accent: #fbbf24;       /* 강조 */
--ok: #10b981;
--bad: #ef4444;
--warn: #f59e0b;
```

### 점수/등급 컬러 스케일 (5단계일 때)

```css
--score-5: #059669;  /* 최고 */
--score-4: #10b981;
--score-3: #f59e0b;
--score-2: #f97316;
--score-1: #ef4444;  /* 최악 */
```

### 타이포그래피

- 폰트: Pretendard Variable (CDN)
- Fallback: -apple-system, "Apple SD Gothic Neo", "Malgun Gothic"
- h1: 28px / 700
- 섹션 제목: 16px / 700
- 본문: 14px / 400
- KPI 숫자: 32px / 700 / letter-spacing: -0.02em

### 레이아웃

- max-width: 1400px (container)
- 카드 radius: 14px
- 카드 padding: 20~24px
- shadow: `0 1px 2px rgba(15,23,42,.04), 0 4px 16px rgba(15,23,42,.05)`
- 반응형: 960px(태블릿 2열→1열), 640px(모바일 1열)

## 섹션 구성 표준 순서

1. **Header** — 제목 + eyebrow 태그 + 메타(원본/생성일/응답수)
2. **KPI Grid** — 4~5개 카드 (auto-fit minmax 200px)
3. **Charts** — 2열 그리드 차트 2개 (분포 + 범주별) + 풀폭 차트 1개 (추이)
4. **Feedback Cards** — text 컬럼 있을 시 3열 (긍정/부정/제안 or 의미 분류)
5. **Data Table** — 필터·검색·정렬·페이징 + 점수·범주 배지
6. **Footer** — 크레딧

## 인터랙션 요구사항

- **KPI 카드**: 호버 시 살짝 떠오르는 transform
- **차트**: Chart.js 4.x, 툴팁 한국어, 반응형
- **테이블**:
  - 컬럼 헤더 클릭 → 정렬 (asc/desc 토글, 화살표 표시)
  - 드롭다운 필터 (범주형 컬럼 자동)
  - 검색창 (text 컬럼 대상, 즉시 반응)
  - 페이지네이션 (기본 10개, 윈도우 5)
  - 초기화 버튼
- **피드백 카드**: 스크롤 영역, 좌측 컬러 바 (긍정:초록, 부정:빨강, 제안:파랑)

## 예시 프롬프트 → 동작

**예시 1: 설문 데이터**
> "설문 CSV 올렸어. 대시보드로 만들어줘."
→ 프로파일링 → numeric(만족도)·categorical(부서)·date(교육일)·text(의견) 감지 → 계획 제시 → 승인 대기 → 생성

**예시 2: 매출 데이터**
> "월별 매출 엑셀이야. 시각화해줘."
→ numeric(매출)·categorical(제품군)·date(월) 감지 → KPI(총매출/월평균/최고월) + 차트(월 추이/제품별) → 테이블

**예시 3: HR 데이터**
> "인사 데이터 대시보드."
→ categorical(부서/직급)·numeric(연차/평가)·date(입사일) → KPI(인원/평균 연차) + 분포 차트 + 필터 테이블

## 주의사항

- 개인정보성 컬럼(이메일·전화·주민번호)은 자동 마스킹 (`xxx@***.com`)
- 응답자 이름 등 식별자는 기본적으로 테이블에서 숨김, 사용자가 명시 요청 시만 포함
- 빈 컬럼·결측치는 렌더링 시 `-`로 표시, 집계에서 제외
- 컬럼명이 너무 길면 40자 내로 축약 (원본은 title 속성에 보존)
- 한국어·영어 혼용 컬럼 모두 지원

## 산출물 체크리스트

생성 완료 후 반드시 확인:
- [ ] HTML이 유효한가 (DOCTYPE, 닫힘 태그)
- [ ] DATA 배열 JSON 파싱 가능한가
- [ ] Chart.js CDN 로드 확인
- [ ] 모든 플레이스홀더 치환 완료 (`__*__` 잔존 없음)
- [ ] 파일 크기 적정 (보통 40~200KB)
- [ ] 원본 파일 건드리지 않았는가
- [ ] output/ 폴더에 저장되었는가
- [ ] `[원본파일명]_dashboard.html` 규칙 준수

## 템플릿·스크립트 위치

- 메인 템플릿: `templates/dashboard_template.html`
- 데이터 프로파일링: `scripts/profile_data.py`
- 데이터 주입: `scripts/inject_data.py`
- 예시 사용 시나리오: `examples/example_usage.md`

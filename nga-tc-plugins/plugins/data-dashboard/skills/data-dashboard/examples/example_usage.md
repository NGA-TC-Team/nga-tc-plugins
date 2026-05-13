# data-dashboard 스킬 사용 예시

## 예시 1. 설문 응답 CSV (카카오스타일 AI 교육 만족도 케이스)

### 입력
```
[카카오스타일] AI 교육 만족도 설문.csv
- 76개 응답
- 컬럼: 타임스탬프, 부서, 교육일, 만족도(1-5), 만족 이유, 불만족 이유, 희망 주제
```

### 사용자 프롬프트
> 해당 csv를 가지고 깔끔한 HTML 대시보드로 만들어줘. 데이터의 핵심 지표를 카드 형태로 상단에 표시하고, 필터/정렬 기능 포함, 반응형, 한국어 폰트 사용, 파일명은 [원본파일명]_dashboard.html로 저장.

### Claude 수행 흐름

**1) 프로파일링**
```bash
python3 scripts/profile_data.py "[카카오스타일] AI 교육 만족도 설문.csv" --out /tmp/profile.json
```

**2) 계획 제시**
```
📊 데이터 개요
- 총 76개 행, 7개 컬럼
- 타입: date(2), numeric(1), categorical(1), text(3)

📌 KPI 카드 (5개)
1. 총 응답 수 76건
2. 평균 만족도 4.72 / 5.0
3. 만족률 98.7% (4점 이상)
4. 참여 부서 3개 (MD 56 · AMD 19 · PO 1)
5. 개선 의견 5건

📈 차트
1. 만족도 분포 (막대, 점수별 색상)
2. 부서별 평균 만족도 (가로 바)
3. 교육일별 추이 (바 + 라인)

💬 피드백 카드
- 만족 이유 / 개선 의견 / 희망 주제

🗂 테이블
- 필터: 부서, 만족도, 교육일
- 검색: 만족 이유·개선 의견·희망 주제
- 정렬: 전체 컬럼 클릭 정렬

승인?
```

**3) 승인 후 생성**
```bash
python3 scripts/inject_data.py /tmp/profile.json templates/dashboard_template.html \
  "output/[카카오스타일] AI 교육 만족도 설문_dashboard.html" \
  --title "카카오스타일 AI 교육 만족도 대시보드" \
  --source "[카카오스타일] AI 교육 만족도 설문.csv" \
  --gen-date "2026-04-22"
```

### 출력
`output/[카카오스타일] AI 교육 만족도 설문_dashboard.html`

---

## 예시 2. 매출 데이터 XLSX

### 입력
```
2026_Q1_매출.xlsx
- 월, 제품군, 매출, 지역 컬럼
```

### Claude 수행

프로파일링 결과:
- date: 월
- categorical: 제품군, 지역
- numeric: 매출

제안 KPI:
- 총 매출 / 월평균 / 최고 매출월 / 제품군 수 / 지역 수

제안 차트:
- 매출 분포 히스토그램
- 제품군별 평균 매출 (가로 바)
- 월별 매출 추이 (라인)

---

## 예시 3. HR 평가 데이터

### 입력
```
hr_evaluation_2026.csv
- 이름, 부서, 직급, 입사일, 평가점수, 코멘트
```

### Claude 수행

**자동 마스킹**: "이름" 컬럼은 기본적으로 테이블에서 숨김 (개인정보 보호).

프로파일:
- identifier: 이름 → 제외
- categorical: 부서, 직급
- date: 입사일
- numeric: 평가점수
- text: 코멘트

KPI: 총 인원 / 평균 평가 / 최고 등급 비율 / 부서 수

---

## 커스터마이징 가이드

사용자가 계획 단계에서 "KPI에 '만족률' 꼭 넣어줘" 같은 요청을 하면:

**profile.json의 suggest 구조를 수정한다**
```json
{
  "suggest": {
    "kpis": [
      {"label":"총 응답 수","kind":"count"},
      {"label":"만족률","kind":"ratio","col":"만족도","threshold":4,"unit":"%","accent":"accent"}
    ]
  }
}
```

지원 KPI `kind`:
- `count`: 전체 행 수
- `avg`: numeric 컬럼 평균
- `max` / `min`: 최대/최소
- `sum`: 합계
- `ratio`: threshold 이상 비율 (만족률, 달성률 등)
- `cat_count`: 범주 고유값 수
- `count_filter`: 특정 컬럼에 값이 있는 행 수

지원 차트 `kind`:
- `distribution`: 값별 카운트 막대
- `group_avg`: 그룹별 평균 가로 바
- `trend`: 시계열 추이 (바 + 라인)

---

## 디버깅 팁

- 템플릿에 `__PLACEHOLDER__`가 남아있다면 → `inject_data.py` 치환 누락, 로그 확인
- JSON 파싱 오류 → 데이터에 특수문자 포함 시 `json.dumps(ensure_ascii=False)` 사용
- 차트가 안 그려짐 → 브라우저 콘솔에서 Chart.js 로드 여부 확인
- 모바일 레이아웃 깨짐 → @media 브레이크포인트 점검 (960px / 640px)

---

## 팀 공유 시 주의

- 이 스킬 폴더 전체를 팀원 각자의 `~/.claude/skills/data-dashboard/` 경로에 복사해야 함
- Cowork 데스크탑의 경우: Skills 폴더 설정 위치 확인 후 배포
- `openpyxl`, `python3` 실행 가능 환경 필요 (대부분 기본 포함)

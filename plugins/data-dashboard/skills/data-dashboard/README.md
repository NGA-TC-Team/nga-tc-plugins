# data-dashboard 스킬

CSV 또는 XLSX 파일을 업로드 받아 인터랙티브 HTML 대시보드를 자동 생성하는 NGA 내부 스킬.

## 기능

- 상단 KPI 카드 4~5개 (자동 추출된 핵심 지표)
- Chart.js 기반 차트 2~3종 (분포·범주별 평균·시계열 추이)
- 긍정/개선/제안 피드백 카드 리스트
- 필터·검색·정렬·페이지네이션 가능한 데이터 테이블
- Pretendard 웹폰트 + 반응형 (960/640 브레이크포인트)
- 단일 HTML 파일 (CDN 제외 외부 의존성 없음)

## 팀원 설치 방법

### 방법 1. 폴더 복사 (권장)

각 팀원이 자신의 컴퓨터에서:

**macOS / Linux**
```bash
# Cowork/Claude Code 스킬 경로
mkdir -p ~/.claude/skills
cp -r /path/to/data-dashboard ~/.claude/skills/
```

**Windows**
```powershell
# PowerShell
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.claude\skills"
Copy-Item -Recurse "C:\path\to\data-dashboard" "$env:USERPROFILE\.claude\skills\"
```

### 방법 2. 압축 파일 배포

```bash
# 압축
cd /path/to/output
zip -r data-dashboard.zip data-dashboard/

# 팀원에게 전달 → 팀원은 ~/.claude/skills/ 에 해제
```

## 설치 확인

1. Cowork 또는 Claude Code 재시작
2. 새 세션에서 "사용 가능한 스킬"에 `data-dashboard` 노출 확인
3. 테스트:
   > "샘플 CSV로 대시보드 만들어줘"

## 사용법

### 기본 사용
1. 설문/매출/HR 등 CSV·XLSX 파일을 워크스페이스에 업로드
2. 프롬프트: "이 파일로 대시보드 만들어줘" 또는 "HTML 대시보드로 시각화해줘"
3. Claude가 데이터 프로파일링 후 KPI·차트 구성안을 제시
4. 승인 후 `output/[원본파일명]_dashboard.html` 생성

### 커스텀 요청
- "KPI에 월별 성장률 추가해줘"
- "부서 컬럼은 테이블에서 숨겨줘" (개인정보 보호)
- "만족률 지표 넣고 4점 이상 기준으로 계산해줘"
- "차트 색상을 브랜드 컬러로 바꿔줘"

## 폴더 구조

```
data-dashboard/
├── SKILL.md                          # 메인 스킬 가이드
├── README.md                         # (이 파일) 설치·사용 안내
├── templates/
│   └── dashboard_template.html       # HTML 템플릿 (플레이스홀더 포함)
├── scripts/
│   ├── profile_data.py               # CSV/XLSX 프로파일링
│   └── inject_data.py                # 데이터 주입·HTML 생성
└── examples/
    └── example_usage.md              # 실제 사용 시나리오
```

## 기술 스택

- **데이터 처리**: Python 3 (표준 라이브러리 + openpyxl)
- **시각화**: Chart.js 4.4 (CDN)
- **폰트**: Pretendard Variable (CDN)
- **스타일**: CSS3 + CSS Grid + CSS Variables
- **인터랙션**: Vanilla JS (외부 프레임워크 없음)

## 주의사항

1. **원본 파일 보존**: 스킬은 원본 CSV/XLSX를 수정/삭제하지 않음
2. **덮어쓰기 방지**: 같은 파일명 존재 시 `_v2`, `_v3` 자동 부여
3. **개인정보**: 이름/이메일 등 identifier 컬럼은 자동 마스킹 또는 테이블에서 숨김
4. **인터넷 필요**: CDN (Chart.js, Pretendard) 로드용

## 업데이트 내역

- **v1.0** (2026-04-22) — 최초 릴리즈, 카카오스타일 AI 교육 만족도 대시보드 케이스 기반

## 문의

NGA 내부 스킬 · 이슈 제보: 잭 (jack@nextgenai.kr)

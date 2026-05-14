# aiq-leader-eval

NGA AI-Q 리더 평가 가이드 v7 HTML을 부서별로 자동 생성하는 Claude Code 스킬.

자가진단 설문 응답(xlsx/csv/tsv)을 받아 잭(NGA)이 경영관리 11명 평가지에서 확정한 표준 양식을 그대로 재현한다.

## 무엇을 만드는가

- 단일 HTML 파일 (`{부서명}_AIQ_리더평가가이드.html`)
- KPI 카드 5종 + 6패널 대시보드 (Chart.js 도넛·막대)
- 멤버별 신호 매트릭스(🟢🟡🔴 × 6종)
- 4축 × 15 sub-item 점수 입력 UI + 자동 축평균·총점·등급 계산
- 멤버별 💾 저장 / 🗑 삭제 / 📝 메모 (브라우저 localStorage)
- 📥 엑셀(XLSX) / CSV 내보내기 (SheetJS)

## 사용법

스킬 인보크:

```
/aiq-leader-eval 마케팅팀 평가지 만들어줘 — [마케팅] AI 활용 현황 공유 설문.xlsx
```

또는 직접 호출:

```bash
python scripts/build_aiq_html.py \
  --input  "/path/to/[마케팅] 설문.xlsx" \
  --department "마케팅" \
  --output "/path/to/마케팅_AIQ_리더평가가이드.html"
```

`--output` 생략 시 입력 파일과 같은 폴더에 `{부서명}_AIQ_리더평가가이드.html`로 저장.

## 4축 구조 (v7)

| 축 | 가중치 | Sub-item | Evidence |
|---|---|---|---|
| 01 AI 툴 활용 | 25% | s11/s12/s14(1-3)/s13(1-4) | q1.3/q2.1/q2.3/q2.2 |
| 02 결과물 완성도 | 30% | ⭐사례 카드(점수X) + s21/s22/s23 | q3.1~q3.6 + q4.1/q4.2/q4.3 |
| 03 속도·임팩트 | 30% | s31/s32/s33/s34 | q5.1/q5.2/q5.3/q5.4·q5.5 |
| 04 지식 공유 | 15% | s41/s42/s43/s44 | q6.1/q6.2/q6.3/리더 종합 |

**총점** = (a1·25% + a2·30% + a3·30% + a4·15%) × 20
**등급** S(85+) / A(70~85) / B(55~70) / C(<55)

## 6가지 신호 (정성 패턴)

응답 패턴에서 자동 분류 — 점수 제안이 아니라 리더 판단 보조용:

| 신호 | 근거 | 색상 라벨 예시 |
|---|---|---|
| `sig_tool` | q2.1 | 🟢 폭넓음 · 🟡 복수 · 🔴 단일 |
| `sig_freq` | q1.3 | 🟢 체화/정착 · 🟡 진행 · 🔴 간헐 |
| `sig_case` | q3.1/q3.3/q3.5 + 길이 | 🟢 풍부 · 🟡 보유/단편 · 🔴 부재 |
| `sig_review` | q4.1/q4.2/q4.3 | 🟢 체계 · 🟡 관리/의식 · 🔴 부재 |
| `sig_impact` | q5.1/q5.2/q5.4 | 🟢 뚜렷 · 🟡 실현 · 🔴 미미 |
| `sig_share` | q6.1/q6.2/q6.3 | 🟢 확산 · 🟡 시도 · 🔴 미공유 |

상세 규칙은 [`references/signal_rules.md`](skills/aiq-leader-eval/references/signal_rules.md).

## 폴더 구조

```
aiq-leader-eval/
├── .claude-plugin/plugin.json
├── README.md
└── skills/aiq-leader-eval/
    ├── SKILL.md
    ├── scripts/
    │   └── build_aiq_html.py    # 메인 빌더 (매핑·신호·집계·HTML 전부)
    └── references/
        ├── column_mapping.md     # 한글 헤더 → q1.3 등 표준 키
        ├── signal_rules.md       # 6가지 신호 분류 로직
        └── axis_schema.md        # 4축 × 15 sub-item v7 정의
```

## 검증 부서

같은 빌더로 다음 부서 HTML이 생성됨 (응답자 N명):

- 경영관리 11명 (v7 원본)
- 마케팅 37명
- HR 22명
- 재무 7명
- 퓨처이노베이션 5명

## 변경 이력

- **1.2.0** — v7 양식 (a2 sub-item 재정의, s44 신규, ⭐사례 헤더카드, 멤버별 저장, 6패널 대시보드)
- **1.1.0** — v6 빌더 스크립트·references 정리, 코어비즈 26명 양식 재현
- **1.0.0** — SKILL.md만 있던 초기 버전

## 라이선스

Internal — NGA 사내 및 계약 고객사 환경 전용.

# NGA-TC Plugins

NGA-TC-Team의 Claude Code 플러그인 마켓플레이스. NGA 컨설팅 현장에서 사용하는 표준 스킬·커맨드·에이전트를 패키징한다.

## 등록 방법

Claude Code 안에서:

```bash
/plugin marketplace add NGA-TC-Team/nga-tc-plugins
```

## 설치 방법

전체 설치:
```bash
/plugin install nga-proofread@nga-tc-plugins
/plugin install aiq-leader-eval@nga-tc-plugins
/plugin install data-dashboard@nga-tc-plugins
/plugin install stq-analysis@nga-tc-plugins
/plugin install api-research@nga-tc-plugins
/reload-plugins
```

## 포함된 플러그인

| 이름 | 카테고리 | 설명 | 버전 |
|---|---|---|---|
| `nga-proofread` | writing | 한국어 텍스트 NGA 톤 교정 | 0.1.0 |
| `aiq-leader-eval` | evaluation | AI-Q 리더 평가 가이드 v7 HTML 생성 | 1.2.0 |
| `data-dashboard` | data | CSV/XLSX → 인터랙티브 대시보드 | 1.0.0 |
| `stq-analysis` | consulting | 업무 SQT 5-Layer 분류·스코어링 | 1.0.0 |
| `api-research` | research | 외부 서비스 API 조사 | 1.0.0 |

## 폴더 구조

```
nga-tc-plugins/
├── .claude-plugin/
│   └── marketplace.json
├── README.md
└── plugins/
    ├── nga-proofread/
    ├── aiq-leader-eval/
    ├── data-dashboard/
    ├── stq-analysis/
    └── api-research/
        ├── .claude-plugin/plugin.json
        └── skills/<name>/SKILL.md (+ scripts/, templates/, examples/)
```

## 신규 플러그인 추가 절차

1. `plugins/<name>/` 폴더 생성
2. `.claude-plugin/plugin.json` 작성
3. `skills/<name>/SKILL.md` 작성 (필요 시 `scripts/`, `templates/`, `examples/` 동봉)
4. 루트 `.claude-plugin/marketplace.json`의 `plugins` 배열에 엔트리 추가
5. 커밋·푸시 후 사용자 측에서 `/plugin marketplace update nga-tc-plugins` 실행

## 버전 관리

시맨틱 버저닝 (MAJOR.MINOR.PATCH):
- MAJOR: 호환성 깨지는 변경
- MINOR: 기능 추가
- PATCH: 버그 수정

## 라이선스

Internal use only. NGA 사내 및 계약 고객사 환경 전용.

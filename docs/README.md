# gemma-docu-writer — 문서 파이프라인 개요

KOICA·ODA 결과보고서 작성을 위한 로컬 웹 도구. 한국어 HWPX 와 영문/국문 PPTX 제안서 두 가지 출력 형식을 각각의 특성에 맞는 **다른 파이프라인** 으로 처리한다.

## 두 출력 형식, 두 파이프라인

| 항목 | HWPX (결과보고서) | PPTX (제안서 양식 기반) |
|------|----------------|---------------------|
| 대상 문서 | 공공기관 결과보고서 (수십 페이지 본문) | 제안서·발표자료 (10~30장 슬라이드) |
| 핵심 엔진 | **비전 기반 스타일 추출** (Ollama/Gemini) | **결정론적 템플릿 주입** (md2pptx-template) |
| LLM 역할 | 양식 이미지 → StyleJSON, 여러 MD → 통합 본문 합성 | **없음** (LLM 호출 0회) |
| 외부 API 필요 | 선택 (Ollama 로컬 또는 Gemini) | **없음** |
| 원본 디자인 보존 | 스타일 참조 (폰트·여백·헤딩) | byte-level 완벽 보존 (XML 그대로) |
| 입력 | 계획서 HWP + Work Plan PDF + Wrap Up PDF → 통합 MD | 단일 MD + 양식 PPTX |
| 출력 | 새 HWPX 파일 | 양식을 수정한 PPTX |
| 상세 문서 | [hwpx-vision-mcp.md](hwpx-vision-mcp.md) | [pptx-md2pptx.md](pptx-md2pptx.md) |

## UI 구성 (localhost:5173)

- 상단 탭 **📝 HWPX / 🎨 PPTX** 로 두 파이프라인 전환
- 공통: 좌측 파일 탐색기, 중앙 MD 프리뷰, 우측 로그·결과 패널
- HWPX 탭: 📝 MD 합성 (여러 MD → 통합 MD, LLM 사용) + 🎯 HWPX 생성 (주입 문서 + 양식 문서 조합)
- PPTX 탭: 단일 카드 "MD + 양식 PPTX → 결과 PPTX" (원클릭, 몇 초)

## 공통 흐름

```
         원본 HWP/PDF ──(kordoc)──▶ 여러 MD
                                      │
                                      ▼
                         📝 MD 합성 (LLM, 공통 단계)
                         여러 MD → 통합 MD 한 개
                                      │
                                 통합 MD 한 개
                                      │
                 ┌────────────────────┴────────────────────┐
                 │                                         │
    ┌──▶ ✍ 슬라이드용 글쓰기                                 │
    │            │                                         │
    │            ▼                                         ▼
    │     🎨 PPTX 변환                               🎯 HWPX 생성
    │   (md2pptx 결정론)                          (비전 + LLM 본문)
    │            │                                         │
    │            ▼                                         ▼
    │      결과 PPTX                                   결과 HWPX
    │            │
    │   ─── 이상 감지 ───
    │            │
    │   ┌────────┴────────┐
    │   │                 │
    │   ▼                 ▼
    │ (A) LLM refiner   (B) 사용자 수기 수정
    │  MD 재작성 제안     양식 PPTX 편집 후 재업로드
    │   │                 │
    └───┴─────────────────┘
        (둘 중 하나 또는 병행 → "슬라이드용 글쓰기" 로 돌아감)
```

- **📝 MD 합성 은 HWPX · PPTX 공통 전처리 단계.** LLM 이 여러 원본 MD 를 하나로 합쳐준다 (상단 `📝 MD 합성` 버튼 — 탭 무관). 결과 MD 는 HWPX · PPTX 양쪽 어디서나 사용 가능.
- **✍ 슬라이드용 글쓰기** 는 PPTX 전용 전처리. 통합 MD 가 줄글 위주면 bullet·표로 재구조화하거나, 양식 PPTX 에 맞게 헤딩 구조를 재편한다.
- **이상 감지 시 피드백 루프는 두 갈래로 끝난다**:
  - **(A) LLM refiner 경로** — analyzer 가 표 넘침·셀 클리핑·미매칭·prose 매핑 실패 등을 감지하면 LLM 이 MD 수정안을 제안 → 사용자 승인 → 재변환.
  - **(B) 사용자 수기 수정 경로** — 사용자가 양식 PPTX 를 직접 편집(행 추가, shape 제거, 폰트 조정 등) 해서 재업로드 → 슬라이드용 글쓰기 단계 재실행.
- 두 경로 모두 **사람이 제어** (자동 무한 루프 없음). 실무에서는 간단한 건 (A), 디자인·레이아웃 변경은 (B) 로 분기.

## 디렉토리

```
.
├── backend/              FastAPI (포트 8765)
│   ├── routes/           hwpx / pptx / files / ollama / report
│   └── services/         composer / llm / renderer / mcp_bridge
├── frontend/             Vite + React (포트 5173)
├── doc_mcp/
│   ├── hwpx_vision/      HWPX 파이프라인 (비전 기반, MCP 서버 포함)
│   └── md2pptx/          PPTX 파이프라인 (결정론적, unpack→edit-XML→pack)
├── docs/                 이 문서들
└── _context/             사용자 자료 (gitignored)
```

## 실행

```bash
install.bat   # 최초 1회 — Python venv, npm install
start.bat     # 백엔드+프론트 동시 기동 + 브라우저 오픈
```

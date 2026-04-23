# gemma-docu-writer (HWPX · PPTX)

로컬 PC 에서 동작하는 React + FastAPI 웹앱. 여러 소스 문서(HWP/HWPX/PDF/MD)를 **HWPX 결과보고서** 또는 **PPTX 제안서** 로 변환한다. 두 출력 포맷은 서로 다른 파이프라인을 사용한다:

- **HWPX** — 비전 기반 스타일 추출 + LLM 본문 합성 (Ollama / Gemini)
- **PPTX** — [md2pptx-template](https://github.com/leedonwoo2827-ship-it/md2pptx-template) 벤더링, **결정론적 템플릿 주입** (LLM·API 없음)

> 상세 설명: [docs/README.md](docs/README.md) · [docs/hwpx-vision-mcp.md](docs/hwpx-vision-mcp.md) · [docs/pptx-md2pptx.md](docs/pptx-md2pptx.md)

## 출력 포맷

| 포맷 | 엔진 | 외부 API | 디자인 보존 |
|------|------|---------|-----------|
| `.hwpx` | 비전 + LLM | Ollama 또는 Gemini (선택) | 스타일 참조 (폰트·여백·헤딩) |
| `.pptx` | md2pptx-template | **없음** | byte-level XML 완벽 보존 |

## 구성

- **frontend/** — Vite + React (포트 5173)
- **backend/** — FastAPI (포트 8765)
- **doc_mcp/**
  - `hwpx_vision/` — HWPX 파이프라인 (analyze_style / apply_style / render_hwp / template_inject)
  - `md2pptx/` — PPTX 파이프라인 (cli / pack / md_parser / slide_scanner / mapper / editor / slide_duplicator / slide_remover / qa)
- **docs/** — 설계 문서
- **_context/** — 입력 문서 보관 (gitignored)

## 사전 요구사항

| 항목 | 용도 | 설치 |
|------|------|------|
| Python 3.11+ | 백엔드 / 변환 엔진 | https://www.python.org |
| Node 18+ | 프론트엔드 | https://nodejs.org |
| Ollama (선택) | 로컬 LLM (HWPX 전용) | https://ollama.com + `ollama pull gemma3n:e4b` |
| Gemini API 키 (선택) | 클라우드 LLM (HWPX 전용) | https://aistudio.google.com |
| LibreOffice | HWP / PDF 렌더링 (선택) | https://www.libreoffice.org |
| kordoc (선택) | HWP → MD 고품질 변환 | `doc_mcp/kordoc` 체크아웃 후 `npm install && npm run build` |

> **PPTX 파이프라인은 LLM 없음** — Ollama / Gemini / LibreOffice 모두 불필요.

## 설치 · 실행

```bat
install.bat   # 최초 1회
start.bat     # 백엔드(8765) + 프론트(5173) 동시 기동 + 브라우저 오픈
```

같은 네트워크의 다른 기기에서는 `http://<이PC의IP>:5173`.

## 상단 탭

### 📝 HWPX 탭

1. **설정(⚙)** 에서 Provider 선택 (Ollama 기본 / Gemini 옵션) + 작업 폴더 지정.
2. 좌측 탐색기에서 HWP/PDF **우클릭 → "MD로 변환"** (kordoc 미설치 시 PDF/HWPX 는 자동 폴백).
3. 외부 MD 는 좌측 드롭존에 드래그앤드롭.
4. 템플릿으로 쓸 `.hwpx` 우클릭 → **"🎯 글쓰기 주입 문서로 지정"** 또는 **"📐 양식 문서로 지정"**.
5. 소스 MD 들을 Ctrl+클릭으로 선택 (템플릿 있으면 1개 이상, 없으면 2개 이상).
6. 상단 **📝 MD 합성** → 템플릿 헤딩 구조 기반 결과보고서 MD 초안 자동 생성.
7. 중앙에서 MD 검수 (외부 에디터로 수정 후 ⟳ 새로고침 가능).
8. **🎯 HWPX 생성** → `*_YYYYMMDD_HHmm.hwpx`.

### 🎨 PPTX 탭

1. (선택) 여러 MD 를 하나로 합치고 싶다면 **HWPX 탭으로 가서 📝 MD 합성 실행** — 통합 MD 가 탐색기에 생성됨. MD 합성은 HWPX·PPTX 공통 전처리.
2. PPTX 탭 전환 → 좌측 상단 **🎨 PPTX 변환** 카드.
3. 탐색기에서 `.md` 클릭 → 카드 MD 칸에 자동 등록 (이미 단일 MD 로 준비됐으면 바로 이 단계).
4. 탐색기에서 양식 `.pptx` 클릭 → 카드 PPTX 칸에 자동 등록.
5. **🚀 변환 시작** → 몇 초 후 같은 폴더에 `{md이름}_result_{timestamp}.pptx` 생성.
6. 카드 하단 결과 섹션: 최종 슬라이드 수, 매칭된 표, 삭제된 슬라이드, 미매칭 경고 표시.

옵션:
- **미매칭 슬라이드 유지** — MD 에 없는 원본 슬라이드를 안 지움 (기본 OFF).
- **미리보기만 (dry-run)** — 매핑 계획만 확인, 파일 미생성.

## Provider 전환 (HWPX 전용)

설정에서 **Ollama ↔ Gemini** 토글. Gemini 선택 시 "문서가 Google 로 전송됨" 경고 배너 노출. 대외비 자료는 Ollama 사용.

## 트러블슈팅

- **Ollama 오프라인 배지**: `ollama serve` 실행 확인 (HWPX 사용 시).
- **HWP 렌더링 실패**: LibreOffice 미설치. 참조 문서를 PDF 로 변환해 `_context/` 에 넣고 재시도.
- **kordoc 없음**: PDF/HWPX 는 자동 폴백 변환됨. HWP(구버전) 는 한/글에서 PDF/HWPX 로 저장 후 업로드.
- **HWPX 변환 중 `python-hwpx` 오류**: `pip install --upgrade python-hwpx` 후 `install.bat` 재실행.
- **PPTX 표가 매핑 안 됨**: MD 표 헤더와 양식 표 헤더의 유사도가 낮음. MD 헤더를 양식에 맞추거나 양식 쪽 열을 MD 에 맞게 수정.
- **PPTX 섹션 H2 복제 실패**: 양식에 "짧은 텍스트 1개인 section-divider 슬라이드" 가 없음. 양식에 1장 추가.

## 배포

- `install.bat` + 소스만 GitHub 업로드.
- `backend/.venv`, `frontend/node_modules`, `doc_mcp/kordoc/node_modules`, `.style_cache/`, `_context/` 의 사용자 자료 (`*.hwp/*.hwpx/*.pdf/*.docx/*.md/*.pptx/pptx_projects/`) 는 `.gitignore` 처리됨.
- LibreOffice / 한컴글꼴 / Ollama 모델은 용량 문제로 별도 배포 권장.

## 아키텍처

```
┌──────────────┐   HTTP   ┌──────────────┐           ┌────────────────────┐
│ React (5173) │ ───────> │ FastAPI(8765)│           │ HWPX: doc_mcp/     │
└──────────────┘          │  /api/...    │ ────────> │  hwpx_vision/      │
                          │              │           │  (비전 + LLM)       │
                          │  composer    │           │                    │
                          │  llm.py ───> │ Ollama    │ PPTX: doc_mcp/     │
                          │  (HWPX only) │ /Gemini   │  md2pptx/          │
                          │              │ ────────> │  (결정론, lxml)     │
                          └──────────────┘           └────────────────────┘
```

## 한계

- **HWPX**: Gemma3n/4n 의 폰트명·pt 정밀도 낮음 → StyleJSON 은 상대 규칙 + 프리셋 병합. "초안 자동 + 사용자 검수" 전제.
- **PPTX**: MD 표 헤더가 양식 표 헤더와 유사하지 않으면 매핑 skip (LLM 이 강제 주입하지 않음). 양식은 "section-divider 스타일 슬라이드 1장 이상" 조건 필요. 이미지·차트·SmartArt 는 원본 유지.

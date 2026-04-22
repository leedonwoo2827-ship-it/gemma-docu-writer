# 결과보고서 작성툴 (HWPX · PPTX)

로컬 PC에서 동작하는 React + FastAPI 웹앱. 여러 소스 문서(HWP/HWPX/PDF/MD)를 LLM으로 합성해 **한국어 결과보고서를 HWPX 또는 PPTX 형식으로 생성**한다. 비전 모델(Gemma4n)이 참조 문서 구조를 분석하고, 템플릿 파일(`.hwpx` / `.pptx`) 의 서식을 유지하면서 본문만 교체하는 방식.

## 출력 포맷

| 포맷 | 생성 경로 | 비고 |
|------|----------|------|
| **`.hwpx`** (한/글) | 🎯 HWPX 생성 버튼 | 템플릿 HWPX 있으면 본문/표 주입, 없으면 단순 변환 |
| **`.pptx`** (PowerPoint/한쇼) | 🎨 PPTX 생성 버튼 | 템플릿 PPTX 필수. 슬라이드별 제목 매칭 후 본문/표 셀 교체 |

> 워드(`.docx`) 포맷은 추후 `doc_mcp/docx_vision/` 으로 추가 예정.

## 구성

- **frontend/** — Vite + React (포트 5173)
- **backend/** — FastAPI (포트 8765)
- **doc_mcp/** — 문서 포맷별 MCP 서버 모음
  - `hwpx_vision/` — HWPX 템플릿 주입 (analyze_style / apply_style / render_hwp / template_inject)
  - `pptx_vision/` — PPTX 템플릿 주입 (list_slides / inject_to_template)
- **_context/** — 입력 문서 보관 (HWP/PDF/MD)

## 사전 요구사항

| 항목 | 용도 | 설치 |
|------|------|------|
| Python 3.11+ | 백엔드/MCP | https://www.python.org |
| Node 18+ | 프론트엔드 | https://nodejs.org |
| Ollama | 로컬 LLM | https://ollama.com |
| `gemma4:e4b` | 비전+텍스트 | `ollama pull gemma4:e4b` |
| LibreOffice | HWP→이미지 (선택) | https://www.libreoffice.org |
| kordoc | HWP→MD (선택) | `doc_mcp/kordoc` 체크아웃 후 `npm install && npm run build` |

> 설정 모달에서 모델 이름을 실제 `ollama list` 결과에 맞게 수정할 수 있다 (자유 입력 지원).

## 설치

```bat
install.bat
```

## 실행

```bat
start.bat
```

브라우저에서 `http://localhost:5173` (같은 네트워크의 다른 기기는 `http://<이PC의IP>:5173`).

## 사용 흐름

1. **설정(⚙)** 에서 작업 폴더 확인, Provider 선택(Ollama 기본 / Gemini 옵션), 모델명 확인.
2. 좌측 탐색기에서 HWP/PDF **우클릭 → "MD로 변환"** (kordoc 없어도 PDF/HWPX는 자동 폴백 변환).
3. **외부에서 만든 MD**라면 좌측 상단 드롭존에 드래그앤드롭.
4. 템플릿으로 쓸 **`.hwpx`** 또는 **`.pptx`** 우클릭 → **"★ 템플릿으로 지정"** → 좌측 하단 "주입 문서" 박스에 고정됨.
5. 소스 MD들을 Ctrl+클릭으로 선택 (템플릿 있으면 1개 이상, 없으면 2개 이상).
6. 상단 **📝 MD 합성** 클릭 → 템플릿 헤딩 구조를 따른 결과보고서 MD 초안 자동 생성 → 중앙에 표시.
7. 중앙 MD 검수 (외부 에디터 수정 후 ⟳ 새로고침 가능).
8. 출력 포맷 선택:
   - **🎯 HWPX 생성** (템플릿 .hwpx 지정 시) → `*_YYYYMMDD_HHmm.hwpx`
   - **🎨 PPTX 생성** (템플릿 .pptx 지정 시) → `*_YYYYMMDD_HHmm.pptx`

## Provider 전환

설정에서 **Ollama ↔ Gemini** 토글. Gemini는 무료 API 키 필요 (https://aistudio.google.com). 선택 시 "문서가 Google로 전송됨" 경고 배너가 표시된다. 대외비 자료는 Ollama 사용.

## 트러블슈팅

- **Ollama 오프라인 배지**: `ollama serve` 실행 확인.
- **`gemma4:e4b` 없음**: `ollama pull gemma4:e4b`.
- **HWP 렌더링 실패**: LibreOffice 미설치. 참조 문서를 PDF로 변환해 `_context/`에 넣고 다시 시도.
- **kordoc 없음**: PDF/HWPX는 자동 폴백 변환됨. HWP(구버전)는 한/글에서 PDF/HWPX로 저장 후 올릴 것.
- **HWPX 변환 중 `python-hwpx` 오류**: `pip install --upgrade python-hwpx` 후 `install.bat` 재실행.
- **주입 결과 "0 섹션 교체"**: MD 헤딩과 템플릿 헤딩이 매칭되지 않음. `📝 MD 합성`으로 템플릿 구조에 맞는 MD를 먼저 생성해서 쓰세요.
- **PPTX 표가 이상하게 교체됨**: LLM이 `|` 구분 행 포맷으로 생성해야 함. 본문 재생성 후 재시도.

## 배포

- `install.bat` + 소스만 GitHub에 업로드.
- `backend/.venv`, `doc_mcp/hwpx_vision/.venv`, `frontend/node_modules`, `doc_mcp/kordoc/node_modules`, `.style_cache/`, `*.hwpx`, `*.pptx` 는 `.gitignore` 처리됨.
- LibreOffice/한컴글꼴/Ollama 모델은 용량 문제로 별도 배포 권장.

## 아키텍처

```
┌──────────────┐   HTTP   ┌──────────────┐  in-proc  ┌────────────────────┐
│ React (5173) │ ───────> │ FastAPI(8765)│ ────────> │ doc_mcp/*_vision   │
└──────────────┘          │  /api/...    │           │  analyze/apply/    │
                          │              │           │  render/inject     │
                          │  composer    │  HTTP     │                    │
                          │  llm.py ───> │ Ollama    │ python-hwpx,       │
                          │              │ /Gemini   │ python-pptx,       │
                          │              │           │ LibreOffice+PyMuPDF│
                          └──────────────┘           └────────────────────┘
```

## 한계

- Gemma4n E4B는 글자 크기 상대 비교/헤딩 계층 추출에는 쓸만하지만 폰트명/pt 정밀도는 낮다 → StyleJSON은 상대 규칙 + 프리셋 병합.
- 템플릿 주입은 **본문 텍스트만 교체** — 이미지/차트/애니메이션/SmartArt는 원본 유지.
- "완전 자동"이 아닌 "초안 자동 + 사용자 검수" 파이프라인으로 설계됨.

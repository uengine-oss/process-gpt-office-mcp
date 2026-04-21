## process-gpt-office-mcp

HWPX·DOCX·슬라이드 문서를 LLM으로 자동 생성·편집하는 **FastMCP HTTP 서버**입니다.  
템플릿 채움, 페이지 단위 HTML 편집, Supabase 업로드까지 하나의 MCP 서버로 처리합니다.

--- 

## 주요 기능

| 영역 | 설명 |
|------|------|
| **HWPX** | 템플릿 URL 다운로드 → LLM 내용 채움 → 결과 HWPX·HTML 생성 |
| **DOCX** | 템플릿 스키마 추출 → LLM 채움 → DOCX·HTML 생성 |
| **슬라이드** | 마크다운 리포트 또는 리서치 목표 → 슬라이드 마크다운·이미지 URL 생성 |
| **페이지 편집** | HWPX/DOCX를 HTML로 변환 후 페이지 단위 수정 → 원본 포맷에 반영 저장 |
| **RAG 연동** | `process-gpt-memento`로 테넌트별 문서 검색·이미지 검색 |
| **이미지 생성** | Google Gemini를 통한 이미지 생성·편집 (BinData 삽입 포함) |
| **웹 리서치** | Tavily API로 슬라이드 작성용 실시간 웹 검색 |

---

## 아키텍처

```
                    ┌─────────────────────────────────┐
                    │         main.py  :1192           │
                    │                                  │
                    │  FastMCP HTTP  │  REST API       │
                    │  /mcp/         │  /api/*         │
                    └───────┬────────┴────────┬────────┘
                            │                 │
              ┌─────────────▼──────────────┐  │ 이미지 편집/개선
              │      office_mcp/           │  │ (뷰어 연동)
              │                            │  │
              │  mcp_server.py (8 도구)    │  │
              │  ├── formats/hwpx/         │  │
              │  ├── formats/docx/         │  │
              │  ├── formats/slides/       │  │
              │  ├── core/ (파싱·청킹·채움)│  │
              │  ├── agent/ (LLM 호출)     │  │
              │  ├── images.py (Gemini)    │  │
              │  └── memento.py (RAG)      │  │
              └────────────────────────────┘  │
                         │                    │
       ┌─────────────────┼─────────────────────────────────┐
       │                 │                   │              │
  OpenAI /          Google Gemini       Supabase       process-gpt-memento
  Gemini LLM        이미지 생성·편집    파일 스토리지   (RAG / 이미지 검색)
                                                        + Tavily (웹 검색)
```

---

## MCP 도구 목록

| 도구 | 설명 |
|------|------|
| `list_reference_documents` | memento에 등록된 참고 문서·폴더 목록 조회 |
| `generate_hwpx` | 주제·설명·참고 문서 기반으로 HWPX 템플릿을 LLM이 채워 결과 파일 생성 |
| `save_hwpx_from_html` | 편집된 HTML(`data-id` 기반)을 원본 HWPX에 반영 후 재저장 |
| `edit_hwpx_page_html` | HWPX를 HTML로 변환하여 특정 페이지 편집 제안 생성 |
| `generate_docx` | DOCX 템플릿 스키마 추출 → LLM 채움 → 결과 DOCX·HTML 생성 |
| `edit_docx_page_html` | DOCX를 HTML로 변환하여 특정 페이지 편집 제안 생성 |
| `save_docx_from_html` | 편집된 HTML을 원본 DOCX에 반영 후 재저장 |
| `generate_slides` | 리포트 마크다운 또는 리서치 목표로부터 슬라이드 마크다운·이미지 URL 생성 |

### REST API (뷰어 연동)

| 엔드포인트 | 설명 |
|------------|------|
| `POST /api/edit-slide-image` | 이미지 URL을 받아 Gemini로 편집 후 base64 반환 |
| `POST /api/enhance-image` | base64 이미지를 AI로 개선 후 base64 반환 |

---

## 실행 방법

### 로컬 실행

```bash
pip install -r requirements.txt
playwright install chromium   # Playwright 브라우저 설치 (첫 실행 시)

python main.py
```

기본 포트: **`1192`**  
MCP 엔드포인트: `http://localhost:1192/mcp/`


---

## MCP 클라이언트 연결 예시

```json
{
  "mcpServers": {
    "office-mcp": {
      "url": "http://localhost:1192/mcp/",
      "transport": "http"
    }
  }
}
```

---

## 의존성

| 패키지 | 용도 |
|--------|------|
| `fastmcp` | MCP HTTP 서버 프레임워크 |
| `openai` | OpenAI LLM 호출 |
| `google-genai` | Google Gemini LLM·이미지 생성·편집 |
| `supabase` | 결과 파일 스토리지 업로드 |
| `python-docx` | DOCX 파일 생성·편집 |
| `playwright` | HWPX HTML 청크 스크린샷 (비전 분석용) |
| `Pillow` | 이미지 처리 |
| `requests` | HTTP 클라이언트 |

---
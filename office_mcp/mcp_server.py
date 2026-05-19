import asyncio
import logging
import os
import re
import shutil
import tempfile
from datetime import datetime
import uuid
from pathlib import Path
from typing import Annotated, Optional
from urllib.parse import quote, urlparse
from html import unescape

import requests
from fastmcp import FastMCP
from pydantic import Field
from supabase import create_client

from .config import DEBUG_OUTPUT_DIR, DEBUG_OUTPUT_ENABLED, IMAGE_GENERATION_ENABLED, LOG_PATH
from .formats.hwpx.runner import process_hwpx_file
from .formats.hwpx.hwpx_to_html import hwpx_to_html
from .formats.hwpx.hwpx_edit import apply_html_edits_to_hwpx
from .core.html_pages import extract_pages, extract_first_page
from .core.html_edit import extract_fills_and_ids
from .agent.agent import _call_llm_text, _call_llm_json
from .formats.docx.template import (
    extract_template_schema,
    apply_schema_output,
    load_template_schema_from_url,
    upload_docx_to_storage,
    _download_docx,
)
from .formats.docx.generation import build_docx_output_from_schema
from .formats.docx.docx_to_html import docx_to_html as _docx_to_html
from .formats.docx.docx_edit import apply_html_edits_to_docx as _apply_html_edits_to_docx
from .formats.slides.generation import (
    build_slide_markdown,
    build_slide_markdown_from_research,
    parse_slides,
    build_style_guide,
    generate_slide_images,
)


class _SuppressLlmFilter(logging.Filter):
    """stream handler에서 LLM 관련 로그(요청/응답/raw/reasoning 등)를 숨긴다. 파일에는 그대로 기록."""
    def filter(self, record: logging.LogRecord) -> bool:
        return "[LLM " not in record.getMessage()


def _setup_logging() -> logging.Logger:
    """office_mcp 패키지 루트 로거에 핸들러를 부착한다.

    각 모듈은 `logger = logging.getLogger(__name__)` 관용구를 쓰면 자식 로거가
    자동으로 이 루트로 propagate되어 파일/스트림에 기록된다. 이름 하드코드가
    필요 없고, 한 곳에서만 핸들러를 관리한다.
    """
    log_path = Path(LOG_PATH)
    if log_path.exists():
        log_path.unlink()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    pkg_logger = logging.getLogger("office_mcp")
    pkg_logger.setLevel(logging.INFO)
    pkg_logger.handlers.clear()
    pkg_logger.propagate = False  # 서드파티 로그(httpx 등)와 격리

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s - %(message)s"
    )
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(fmt)
    stream_handler.addFilter(_SuppressLlmFilter())
    pkg_logger.addHandler(file_handler)
    pkg_logger.addHandler(stream_handler)
    return logging.getLogger(__name__)


logger = _setup_logging()
DEBUG_OUTPUT_PATH = Path(__file__).resolve().parent / DEBUG_OUTPUT_DIR
from .config import SUPABASE_URL, SUPABASE_KEY
SUPABASE_BUCKET = "deep_research_files"
HWPX_CONTENT_TYPE = "application/vnd.hancom.hwpx"
HTML_CONTENT_TYPE = "text/html; charset=utf-8"

mcp = FastMCP("process-gpt-office-mcp")


def _safe_filename_from_url(url: str) -> str:
    name = Path(urlparse(url).path).name
    if not name:
        return "template.hwpx"
    return name


def _safe_storage_name(filename: str) -> str:
    raw = (filename or "output.hwpx").strip()
    safe = re.sub(r"[^0-9A-Za-z._-]+", "_", raw).strip("_")
    if not safe:
        safe = "output.hwpx"
    if not safe.lower().endswith(".hwpx"):
        safe += ".hwpx"
    return safe


def _safe_html_name(filename: str) -> str:
    raw = (filename or "output.html").strip()
    safe = re.sub(r"[^0-9A-Za-z._-]+", "_", raw).strip("_")
    if not safe:
        safe = "output.html"
    if not safe.lower().endswith(".html"):
        safe += ".html"
    return safe


def _build_output_basename(report_topic: str) -> str:
    topic = (report_topic or "").strip()
    safe_topic = re.sub(r"[^0-9A-Za-z._-]+", "_", topic).strip("_")
    if not safe_topic:
        safe_topic = "report"
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return f"filled-{safe_topic}_{stamp}"


def _build_edit_basename(filename: str) -> str:
    safe = _safe_storage_name(filename or "output.hwpx").replace(".hwpx", "")
    stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
    return f"edited-{safe}_{stamp}"


def _build_page_edit_prompt(original_page_html: str, instruction: str) -> tuple[str, str]:
    prompt_sys = (
        "당신은 HWPX 문서를 편집하는 전문가입니다. "
        "아래 HTML은 단일 페이지를 나타내며 data-id를 포함합니다. "
        "구조와 data-id를 유지하고 텍스트만 수정하세요. "
        "페이지 외부 내용은 절대 변경하지 마세요."
    )
    prompt_user = f"""## 사용자 지시
{instruction}

## 페이지 HTML (수정 대상)
{original_page_html}

## 출력 규칙
1) 반드시 단일 <div class="page"> ... </div> 형태로만 출력
2) data-id 속성을 삭제/변경하지 말 것
3) 표 구조/셀 구조 유지, 텍스트만 수정
4) 마크다운 코드블록 금지, 순수 HTML만 출력
"""
    return prompt_sys, prompt_user


def _build_page_edit_patch_prompt(
    original_page_html: str,
    instruction: str,
) -> tuple[str, str]:
    prompt_sys = (
        "당신은 HWPX 문서를 편집하는 전문가입니다. "
        "아래 HTML은 단일 페이지를 나타내며 data-id를 포함합니다. "
        "구조와 data-id를 유지하고 텍스트만 수정하세요. "
        "페이지 외부 내용은 절대 변경하지 마세요."
    )
    prompt_user = f"""## 사용자 지시
{instruction}

## 페이지 HTML (수정 대상)
{original_page_html}

## 출력(JSON)
{{"edits":[{{"label":"1) 활용 오픈소스 AI(모델)명","new_text":"..."}}]}}

규칙:
1) data-id가 있는 요소만 수정 대상
2) id는 숫자만 사용
3) new_text는 순수 텍스트만 허용 (HTML 태그 금지)
4) 항목명(라벨) 셀은 수정 금지, 값/내용 셀만 수정
5) 가능하면 id 대신 label(라벨 텍스트)을 사용해 지정할 것
6) label은 페이지 내 실제 항목명 텍스트와 일치해야 함
7) HTML을 다시 생성하지 말고 edits만 반환
8) 사용자 지시에 id=숫자 형태가 포함되면 반드시 해당 id를 사용
"""
    return prompt_sys, prompt_user


def _extract_td_rows(page_html: str) -> list[list[int]]:
    rows: list[list[int]] = []
    for row_match in re.finditer(r"<tr[^>]*>(.*?)</tr>", page_html, re.DOTALL):
        row_html = row_match.group(1)
        ids: list[int] = []
        for match in re.finditer(r"<td[^>]*\bdata-id=\"(\d+)\"", row_html):
            try:
                ids.append(int(match.group(1)))
            except (TypeError, ValueError):
                continue
        if ids:
            rows.append(ids)
    return rows


def _normalize_label_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _normalize_patch_text(value: str) -> str:
    text = value or ""
    if "<" in text and ">" in text:
        text = re.sub(r"(?i)<br\s*/?>", "\n", text)
        text = re.sub(r"<[^>]+>", "", text)
    text = unescape(text)
    return text.replace("\xa0", " ").strip()


def _extract_public_url(response: object) -> Optional[str]:
    if not response:
        return None
    if isinstance(response, dict):
        if response.get("publicUrl"):
            return response.get("publicUrl")
        if response.get("public_url"):
            return response.get("public_url")
        data = response.get("data")
        if isinstance(data, dict) and data.get("publicUrl"):
            return data.get("publicUrl")
    return None


def _upload_hwpx_to_storage(file_path: Path, output_name: str) -> str:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("SUPABASE_URL 또는 SUPABASE_KEY가 설정되지 않았습니다.")
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    safe_name = _safe_storage_name(output_name)
    storage_path = f"hwpx/{uuid.uuid4().hex}_{safe_name}"
    file_bytes = file_path.read_bytes()
    resp = supabase.storage.from_(SUPABASE_BUCKET).upload(
        storage_path,
        file_bytes,
        {"content-type": HWPX_CONTENT_TYPE, "upsert": "true"},
    )
    if hasattr(resp, "path") and not resp.path:
        raise RuntimeError(f"storage 업로드 실패: 응답 path 없음 {resp}")
    public = supabase.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)
    url = _extract_public_url(public)
    if url:
        return url
    base_url = SUPABASE_URL.rstrip("/")
    return f"{base_url}/storage/v1/object/public/{SUPABASE_BUCKET}/{quote(storage_path, safe='/-_.')}"


def _upload_html_to_storage(file_path: Path, output_name: str) -> str:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("SUPABASE_URL 또는 SUPABASE_KEY가 설정되지 않았습니다.")
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
    safe_name = _safe_html_name(output_name)
    storage_path = f"hwpx_html/{uuid.uuid4().hex}_{safe_name}"
    file_bytes = file_path.read_bytes()
    resp = supabase.storage.from_(SUPABASE_BUCKET).upload(
        storage_path,
        file_bytes,
        {
            "content-type": "text/html",
            "cache-control": "3600",
            "content-disposition": "inline",
            "upsert": "true",
        },
    )
    if hasattr(resp, "path") and not resp.path:
        raise RuntimeError(f"storage 업로드 실패: 응답 path 없음 {resp}")
    public = supabase.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)
    url = _extract_public_url(public)
    if url:
        return url
    base_url = SUPABASE_URL.rstrip("/")
    return f"{base_url}/storage/v1/object/public/{SUPABASE_BUCKET}/{quote(storage_path, safe='/-_.')}"




# NOTE: 참고문서 사용자 선택 플로우 제거로 툴 노출 비활성화.
#       다시 쓰려면 아래 @mcp.tool 주석만 해제하면 된다.
# @mcp.tool
async def list_reference_documents(
    query: Annotated[str, Field(description="검색 쿼리 (예: 보고서 주제)")],
    tenant_id: Annotated[str, Field(description="테넌트 ID")],
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
) -> dict:
    """memento에서 참고 가능한 문서 목록을 검색하여 반환한다.

    [필수 선행 도구] generate_hwpx / generate_docx 호출 전에 반드시 이 도구를 먼저 호출해야 한다.
    사용자가 HWPX/DOCX 파일 작성을 요청하면 이 도구를 즉시 호출하고, 반환된 JSON을 그대로 응답에 포함하여 반환한다.
    사용자가 문서를 선택한 후에만 generate_hwpx / generate_docx를 호출한다.
    반환된 문서 목록을 사용자에게 보여주고, 사용자가 선택한 문서명 리스트를
    generate_hwpx 또는 generate_docx의 reference_documents 파라미터로 전달한다.
    """
    if not tenant_id or not tenant_id.strip():
        raise ValueError("tenant_id is required")
    if not query or not query.strip():
        raise ValueError("query is required")

    tenant_id = tenant_id.strip()
    try:
        from .memento import _list_documents_with_folders
        doc_details = await _list_documents_with_folders(tenant_id)
        if not doc_details:
            return {
                "documents": [],
                "total": 0,
                "message": "등록된 문서가 없습니다.",
            }

        # 폴더 경로에서 마지막 세그먼트(실제 폴더명)만 추출
        # 예: "localhost/deep_research_source/유엔진_기업정보" → "유엔진_기업정보"
        def _leaf_folder(folder_name: str) -> str:
            if not folder_name:
                return ""
            parts = [p for p in folder_name.split("/") if p.strip()]
            if not parts:
                return ""
            return parts[-1]

        return {
            "documents": doc_details,
            "total": len(doc_details),
            "message": f"{len(doc_details)}개의 문서를 찾았습니다. 사용자에게 어떤 문서를 참고할지 선택하게 해주세요.",
            "user_request_type": "select_items",
            "question": "제안서 작성에 참고할 문서를 선택해 주세요.",
            "items": [
                {
                    "id": d["file_name"],
                    "label": d["file_name"],
                    "description": _leaf_folder(d.get("drive_folder_name", "")),
                }
                for d in doc_details
            ],
            "allow_multiple": True,
            "min_select": 1,
            "image_options": [
                {
                    "id": "image_generation",
                    "label": "Gemini 이미지 생성",
                    "description": "AI가 문서 내용에 맞는 이미지를 자동으로 생성하여 삽입합니다",
                    "default": True,
                },
                {
                    "id": "image_reference",
                    "label": "Memento 이미지 참조",
                    "description": "지식 베이스에서 관련 이미지를 검색하여 삽입합니다",
                    "default": True,
                },
            ],
        }
    except Exception as exc:
        logger.warning("list_reference_documents 실패: %s", exc)
        raise ValueError(f"문서 목록 조회 실패: {exc}")


@mcp.tool
async def generate_hwpx(
    template_url: Annotated[str, Field(description="HWPX 템플릿 URL")],
    report_topic: Annotated[str, Field(description="보고서 주제")],
    report_description: Annotated[Optional[str], Field(description="보고서 상세 설명")] = "",
    reference_text: Annotated[Optional[str], Field(description="참고할 텍스트")] = "",
    reference_documents: Annotated[Optional[str], Field(description="참고할 문서명 리스트 (JSON 배열 문자열, 예: '[\"문서1.pdf\", \"문서2.docx\"]'). list_reference_documents에서 사용자가 선택한 문서명을 전달한다.")] = "",
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (memento RAG 검색용)")] = "",
    image_generation_enabled: Annotated[Optional[bool], Field(description="Gemini 이미지 생성 사용 여부. 사용자가 이미지 생성을 원하면 true, 원하지 않으면 false. 미지정 시 서버 기본값 사용.")] = None,
    image_reference_enabled: Annotated[Optional[bool], Field(description="Memento 이미지 참조 사용 여부. 사용자가 지식 베이스 이미지 참조를 원하면 true, 원하지 않으면 false. 미지정 시 서버 기본값 사용.")] = None,
    source_chunks_json: Annotated[Optional[str], Field(description="소스 파일에서 추출한 참고자료 청크 리스트 (JSON 배열). 각 청크는 {file_name, chunk_index, original_text, summary} 구조. 프로세스 소스에서 파싱된 참고자료를 템플릿 채우기에 활용한다.")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
    room_id: Annotated[Optional[str], Field(description="채팅방 ID (자동 주입). 있으면 방에 업로드된 문서만 참조하며 드라이브 검색을 스킵한다.")] = "",
    proc_inst_id: Annotated[Optional[str], Field(description="프로세스 인스턴스 ID. 있으면 해당 프로세스에 ingest된 문서만 참조한다.")] = "",
) -> dict:
    """HWPX 템플릿을 채워 스토리지 URL로 반환한다.

    사용자가 HWPX 작성을 요청하면 즉시 이 도구를 호출한다 (참고문서 선택 플로우 없음).
    reference_documents가 비어 있으면 memento 전체 검색으로 관련 자료를 자동 조회한다.
    사용자가 이미지 생성/참조 옵션을 선택한 경우 image_generation_enabled, image_reference_enabled에 전달한다.
    """
    if not template_url:
        raise ValueError("template_url is required")
    if not report_topic:
        raise ValueError("report_topic is required")

    import json as _json

    report_description = report_description or ""
    reference_text = reference_text or ""

    # reference_documents 파싱 (JSON 배열 문자열 → 리스트)
    selected_docs: list = []
    if reference_documents and reference_documents.strip():
        try:
            parsed = _json.loads(reference_documents.strip())
            if isinstance(parsed, list):
                selected_docs = [str(d).strip() for d in parsed if d]
            elif isinstance(parsed, str):
                selected_docs = [parsed.strip()]
        except _json.JSONDecodeError:
            # JSON이 아니면 콤마 구분으로 시도
            selected_docs = [d.strip() for d in reference_documents.split(",") if d.strip()]
        logger.info("generate_hwpx: 사용자 선택 문서 %d개: %s", len(selected_docs), selected_docs)

    # tenant_id가 있으면 memento에서 내부 지식 자료를 검색해 reference_text에 추가
    if (tenant_id or "").strip():
        try:
            from .memento import search_memento_smart, search_memento_by_documents, sources_to_reference_text
            logger.info("generate_hwpx: memento RAG 검색 시작 (tenant_id=%s)", tenant_id)

            if selected_docs:
                # 사용자가 선택한 문서만 참고
                memento_sources = await search_memento_by_documents(
                    query=report_topic,
                    outline=[report_description] if report_description else [report_topic],
                    tenant_id=tenant_id.strip(),
                    file_names=selected_docs,
                    room_id=(room_id or "").strip() or None,
                    proc_inst_id=(proc_inst_id or "").strip() or None,
                )
                logger.info("generate_hwpx: 선택 문서 기반 검색 완료 (%d개 소스)", len(memento_sources))
            else:
                # room_id > proc_inst_id > (드라이브 스마트는 현재 비활성) 순으로 분기
                memento_sources = await search_memento_smart(
                    query=report_topic,
                    outline=[report_description] if report_description else [report_topic],
                    tenant_id=tenant_id.strip(),
                    room_id=(room_id or "").strip() or None,
                    proc_inst_id=(proc_inst_id or "").strip() or None,
                )

            if memento_sources:
                memento_text = sources_to_reference_text(memento_sources)
                if reference_text.strip():
                    reference_text = memento_text + "\n\n---\n\n" + reference_text
                else:
                    reference_text = memento_text
                logger.info("generate_hwpx: memento 소스 %d개 → reference_text에 추가 완료", len(memento_sources))
            else:
                logger.info("generate_hwpx: memento 검색 결과 없음")
        except Exception as exc:
            logger.warning("generate_hwpx: memento 검색 실패 (무시하고 계속 진행): %s", exc)
    else:
        logger.info("generate_hwpx: tenant_id 없음 → memento 검색 건너뜀")

    # source_chunks_json 파싱 (프로세스 소스에서 온 참고자료 청크)
    source_chunks: list = []
    if source_chunks_json and source_chunks_json.strip():
        try:
            parsed_chunks = _json.loads(source_chunks_json.strip())
            if isinstance(parsed_chunks, list):
                source_chunks = parsed_chunks
                logger.info("generate_hwpx: 소스 참고자료 청크 %d개 수신", len(source_chunks))
        except _json.JSONDecodeError:
            logger.warning("generate_hwpx: source_chunks_json 파싱 실패")

    template_name = _safe_filename_from_url(template_url)
    base_name = _build_output_basename(report_topic)
    output_name = f"{base_name}.hwpx"
    output_html_name = f"{base_name}.html"
    logger.info("generate_hwpx start: template=%s output=%s", template_name, output_name)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        output_path = tmp_dir_path / output_name
        html_output_path = tmp_dir_path / output_html_name

        response = requests.get(template_url, timeout=60)
        response.raise_for_status()
        template_path.write_bytes(response.content)

        await process_hwpx_file(
            str(template_path),
            str(output_path),
            report_topic=report_topic,
            report_description=report_description,
            reference_text=reference_text,
            tenant_id=(tenant_id or "").strip(),
            selected_file_names=selected_docs or None,
            image_generation_enabled=image_generation_enabled,
            image_reference_enabled=image_reference_enabled,
            source_chunks=source_chunks or None,
            proc_inst_id=(proc_inst_id or "").strip() or None,
            room_id=(room_id or "").strip() or None,
        )

        if DEBUG_OUTPUT_ENABLED:
            DEBUG_OUTPUT_PATH.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(output_path, DEBUG_OUTPUT_PATH / output_name)
        file_url = _upload_hwpx_to_storage(output_path, output_name)
        # 출처 사이드카 (runner가 저장함) 로드 — 없어도 무해
        _sources_map = None
        try:
            import json as _json_sc
            _sidecar = Path(str(output_path) + ".sources.json")
            if _sidecar.exists():
                with open(_sidecar, "r", encoding="utf-8") as _fp:
                    _sources_map = _json_sc.load(_fp)
                logger.info("출처 사이드카 로드: %d개 노드", len(_sources_map or {}))
        except Exception as _e:
            logger.warning("출처 사이드카 로드 실패: %s", _e)
        html_url = ""
        try:
            hwpx_to_html(output_path, html_output_path, use_lineseg=False, inject_ids=True, sources_map=_sources_map)
            html_url = _upload_html_to_storage(html_output_path, output_html_name)
        except Exception as e:
            logger.warning("HTML 변환 실패 (HWPX만 반환): %s", e)
            html_url = ""
        if html_url:
            logger.info("HTML 변환 완료: output=%s url=%s", output_html_name, html_url)
        else:
            logger.warning("HTML 변환 실패 (HWPX만 반환)")

    logger.info("generate_hwpx done: output=%s url=%s", output_name, file_url)
    return {
        "file_name": output_name,
        "content_type": HWPX_CONTENT_TYPE,
        "file_url": file_url,
        "html_name": output_html_name if html_url else "",
        "html_content_type": HTML_CONTENT_TYPE if html_url else "",
        "html_url": html_url,
    }


@mcp.tool
async def save_hwpx_from_html(
    hwpx_url: Annotated[str, Field(description="원본 HWPX URL")],
    edited_html: Annotated[str, Field(description="편집된 HTML (data-id 포함)")],
    output_name: Annotated[Optional[str], Field(description="저장할 HWPX 파일명")] = "",
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (자동 주입)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
) -> dict:
    """편집된 HTML을 HWPX로 반영하고 스토리지 URL로 반환한다."""
    if not hwpx_url:
        raise ValueError("hwpx_url is required")
    if not edited_html:
        raise ValueError("edited_html is required")

    template_name = _safe_filename_from_url(hwpx_url)
    base_name = _build_edit_basename(output_name or template_name)
    output_hwpx_name = f"{base_name}.hwpx"
    output_html_name = f"{base_name}.html"
    edited_html_name = f"{base_name}_edited.html"
    logger.info("save_hwpx_from_html start: template=%s output=%s", template_name, output_hwpx_name)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        output_path = tmp_dir_path / output_hwpx_name
        html_output_path = tmp_dir_path / output_html_name
        edited_html_path = tmp_dir_path / edited_html_name

        response = requests.get(hwpx_url, timeout=60)
        response.raise_for_status()
        template_path.write_bytes(response.content)

        edited_html_path.write_text(edited_html, encoding="utf-8")
        apply_html_edits_to_hwpx(str(template_path), str(output_path), edited_html)

        DEBUG_OUTPUT_PATH.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(output_path, DEBUG_OUTPUT_PATH / output_hwpx_name)
        shutil.copyfile(edited_html_path, DEBUG_OUTPUT_PATH / edited_html_name)

        file_url = _upload_hwpx_to_storage(output_path, output_hwpx_name)
        html_url = ""
        try:
            hwpx_to_html(output_path, html_output_path, use_lineseg=False, inject_ids=True)
            shutil.copyfile(html_output_path, DEBUG_OUTPUT_PATH / output_html_name)
            html_url = _upload_html_to_storage(html_output_path, output_html_name)
        except Exception as e:
            logger.warning("HTML 재변환 실패 (HWPX만 반환): %s", e)
            html_url = ""

    logger.info("save_hwpx_from_html done: output=%s url=%s", output_hwpx_name, file_url)
    return {
        "file_name": output_hwpx_name,
        "content_type": HWPX_CONTENT_TYPE,
        "file_url": file_url,
        "html_name": output_html_name if html_url else "",
        "html_content_type": HTML_CONTENT_TYPE if html_url else "",
        "html_url": html_url,
    }


@mcp.tool
async def edit_hwpx_page_html(
    hwpx_url: Annotated[str, Field(description="원본 HWPX URL")],
    page_number: Annotated[int, Field(description="수정할 페이지 번호 (1부터 시작, 필수)")],
    instruction: Annotated[str, Field(description="수정 지시사항 (페이지 내부에서 무엇을 어떻게 바꿀지 명시)")],
    include_original: Annotated[Optional[bool], Field(description="응답에 원본 페이지 HTML 포함 여부")] = False,
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (자동 주입)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
) -> dict:
    """지정한 페이지를 지시사항대로 수정해 edits를 반환한다.

    요구 사항:
    - page_number와 instruction이 모두 필요
    - 지시사항에는 '어떤 내용을 어떻게 수정할지'를 구체적으로 포함

    입력 예시:
    - page_number: 2
    - instruction: "2페이지의 과제추진 필요성 문단에 기대효과를 1문단 추가"
    - instruction: "3페이지 표에서 '담당부서' 값을 'AI전략팀'으로 변경"
    """
    if not hwpx_url:
        raise ValueError("hwpx_url is required")
    if not page_number or page_number < 1:
        raise ValueError("page_number must be >= 1")
    if not instruction:
        raise ValueError("instruction is required")

    template_name = _safe_filename_from_url(hwpx_url)
    logger.info("edit_hwpx_page_html start: template=%s page=%d", template_name, page_number)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        html_output_path = tmp_dir_path / f"page-edit-{template_name}.html"

        response = requests.get(hwpx_url, timeout=60)
        response.raise_for_status()
        template_path.write_bytes(response.content)

        hwpx_to_html(template_path, html_output_path, use_lineseg=False, inject_ids=True)
        html_text = html_output_path.read_text(encoding="utf-8")
        pages = extract_pages(html_text)
        if not pages:
            raise ValueError("페이지를 추출할 수 없습니다.")
        if page_number > len(pages):
            raise ValueError(
                f"page_number 범위 초과: 이 문서는 총 {len(pages)}페이지입니다. "
                f"page_number는 물리적 페이지 번호(1~{len(pages)})를 사용하세요. "
                f"문서 섹션 번호(예: 4.1)와 다릅니다."
            )
        original_page = pages[page_number - 1]

    prompt_sys, prompt_user = _build_page_edit_patch_prompt(original_page, instruction)
    edits_result = await asyncio.to_thread(_call_llm_json, prompt_sys, prompt_user, 0.2)
    if not isinstance(edits_result, dict):
        raise ValueError("LLM 결과가 올바르지 않습니다.")
    edits = edits_result.get("edits", [])
    if not isinstance(edits, list):
        edits = []

    _orig_fills, orig_ids = extract_fills_and_ids(original_page)
    orig_fills = _orig_fills or {}
    td_rows = _extract_td_rows(original_page)
    label_to_value_id: dict[str, int] = {}
    label_id_to_value_id: dict[int, int] = {}
    for row in td_rows:
        for idx, td_id in enumerate(row):
            next_idx = idx + 1
            if next_idx >= len(row):
                continue
            label_text = _normalize_label_text(str(orig_fills.get(td_id, "")))
            if not label_text:
                continue
            value_id = row[next_idx]
            label_to_value_id[label_text] = value_id
            label_id_to_value_id[td_id] = value_id
    normalized_edits = []
    for item in edits:
        if not isinstance(item, dict):
            continue
        raw_id = item.get("id")
        raw_label = item.get("label")
        new_text = item.get("new_text")
        if new_text is None:
            continue
        target_id: Optional[int] = None
        if raw_label:
            label_key = _normalize_label_text(str(raw_label))
            if label_key in label_to_value_id:
                target_id = label_to_value_id[label_key]
        if target_id is None and raw_id is not None:
            try:
                numeric_id = int(raw_id)
            except (TypeError, ValueError):
                continue
            if orig_ids and numeric_id not in orig_ids:
                continue
            target_id = label_id_to_value_id.get(numeric_id, numeric_id)
        if target_id is None:
            continue
        normalized_edits.append(
            {"id": target_id, "new_text": _normalize_patch_text(str(new_text))}
        )

    logger.info("edit_hwpx_page_html done: page=%d", page_number)
    payload = {
        "page_number": page_number,
        "edits": normalized_edits,
    }
    if include_original:
        payload["original_page_html"] = original_page
    return payload


@mcp.tool
async def generate_docx(
    query: Annotated[str, Field(description="사용자 요청 / 리서치 주제")],
    template_url: Annotated[Optional[str], Field(description="DOCX 템플릿의 http(s) URL. 외부 URL 인 경우만 사용. supabase storage 자료는 template_file_id 로.")] = "",
    template_file_id: Annotated[Optional[str], Field(description="DOCX 템플릿의 supabase storage path (= knowledge_files.source_ref). bucket=files 에서 직접 다운로드. URL 변환 불필요.")] = "",
    sources_json: Annotated[Optional[str], Field(description="소스 목록 (JSON 직렬화된 list[dict])")] = "",
    outline_json: Annotated[Optional[str], Field(description="개요 목록 (JSON 직렬화된 list[str])")] = "",
    user_info_json: Annotated[Optional[str], Field(description="작성자 정보 (JSON 직렬화된 list[dict])")] = "",
    image_hints_json: Annotated[Optional[str], Field(description="이미지 힌트 (JSON 직렬화된 list[dict])")] = "",
    output_name: Annotated[Optional[str], Field(description="저장할 파일명 (확장자 포함)")] = "",
    report_id: Annotated[Optional[str], Field(description="스토리지 경로용 리포트 ID")] = "",
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (memento RAG 검색용)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
    room_id: Annotated[Optional[str], Field(description="채팅방 ID (자동 주입). 있으면 방에 업로드된 문서만 참조하며 드라이브 검색을 스킵한다.")] = "",
    proc_inst_id: Annotated[Optional[str], Field(description="프로세스 인스턴스 ID. 있으면 해당 프로세스에 ingest된 문서만 참조한다.")] = "",
    file_ids_json: Annotated[Optional[str], Field(description="사용자가 명시 선택한 자료 file_id 목록 (JSON 직렬화된 list[str]). 비어있지 않으면 그 자료들에만 검색을 제한하며 room_id/proc_inst_id 보다 우선한다.")] = "",
) -> dict:
    """DOCX 템플릿을 LLM으로 채워 스토리지 URL로 반환한다.

    템플릿 입력은 둘 중 하나:
    - ``template_file_id``: supabase storage path (= ``knowledge_files.source_ref``). bucket=files 에서 직접 다운로드.
      클라이언트(deep-agents-temp)가 사용자 선택 자료에서 자동 검출해 주입한다.
    - ``template_url``: 외부 http(s) URL (Google Drive 공유 링크 등). file_id 가 없을 때만 사용.

    참고문서 선택 플로우 없음 — reference_documents가 비어 있으면 memento 검색으로 자동 조회한다.
    """
    if not query:
        raise ValueError("query is required")
    if not (template_url or "").strip() and not (template_file_id or "").strip():
        raise ValueError("template_url or template_file_id is required")

    import json as _json
    import uuid as _uuid
    from datetime import datetime as _dt

    sources = _json.loads(sources_json) if sources_json else []
    outline = _json.loads(outline_json) if outline_json else []

    # 명시 선택 자료 목록 파싱 — 비어있지 않으면 검색을 그 자료에 한정.
    parsed_file_ids: list[str] = []
    if file_ids_json:
        try:
            raw = _json.loads(file_ids_json)
            if isinstance(raw, list):
                parsed_file_ids = [str(x).strip() for x in raw if str(x).strip()]
        except Exception as exc:
            logger.warning("generate_docx: file_ids_json 파싱 실패 (%s) → 무시하고 진행", exc)

    # tenant_id가 있으면 memento에서 내부 지식 자료를 검색해 sources에 추가
    if (tenant_id or "").strip():
        try:
            from .memento import search_memento_smart
            logger.info(
                "generate_docx: memento RAG 검색 시작 (tenant_id=%s, file_ids=%d)",
                tenant_id, len(parsed_file_ids),
            )
            memento_sources = await search_memento_smart(
                query=query,
                outline=outline if outline else [query],
                tenant_id=tenant_id.strip(),
                room_id=(room_id or "").strip() or None,
                proc_inst_id=(proc_inst_id or "").strip() or None,
                file_ids=parsed_file_ids or None,
            )
            if memento_sources:
                sources = memento_sources + sources
                logger.info("generate_docx: memento 소스 %d개 추가 (총 %d개)", len(memento_sources), len(sources))
            else:
                logger.info("generate_docx: memento 검색 결과 없음")
        except Exception as exc:
            logger.warning("generate_docx: memento 검색 실패 (무시하고 계속 진행): %s", exc)
    else:
        logger.info("generate_docx: tenant_id 없음 → memento 검색 건너뜀")
    user_info = _json.loads(user_info_json) if user_info_json else []
    image_hints = _json.loads(image_hints_json) if image_hints_json else []
    rid = report_id or _uuid.uuid4().hex
    stamp = _dt.utcnow().strftime("%Y%m%d_%H%M%S")
    safe_output = output_name or f"report-{stamp}.docx"
    if not safe_output.lower().endswith(".docx"):
        safe_output += ".docx"
    # Supabase storage rejects non-ASCII keys — use ASCII-safe storage key
    import re as _re, unicodedata as _ud
    _ascii_key = _ud.normalize("NFKD", safe_output).encode("ascii", "ignore").decode("ascii")
    _ascii_key = _re.sub(r"[^\w.\-]", "_", _ascii_key).strip("_") or f"report-{stamp}.docx"
    if not _ascii_key.lower().endswith(".docx"):
        _ascii_key += ".docx"
    storage_path = f"deep-research/{rid}/{_ascii_key}"

    template_source = (
        f"file_id={template_file_id}" if (template_file_id or "").strip()
        else f"url={template_url}"
    )
    logger.info("generate_docx start: template=%s output=%s", template_source, safe_output)

    from docx import Document as _Document
    from .formats.docx.template import _download_docx_from_storage

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        # file_id 우선 — supabase storage 에서 직접 다운로드 (URL 변환 없음).
        if (template_file_id or "").strip():
            template_path = _download_docx_from_storage(template_file_id.strip())
        else:
            template_path = _download_docx(template_url)
        doc = _Document(str(template_path))
        schema = extract_template_schema(doc)

        schema_output = await build_docx_output_from_schema(
            query=query,
            outline=outline,
            sources=sources,
            schema=schema,
            user_info=user_info or None,
            image_hints=image_hints or None,
            tenant_id=(tenant_id or "").strip() or None,
            room_id=(room_id or "").strip() or None,
            proc_inst_id=(proc_inst_id or "").strip() or None,
        )

        apply_schema_output(doc, schema, schema_output, report_id=rid)

        output_path = tmp_dir_path / safe_output
        doc.save(str(output_path))
        file_url = upload_docx_to_storage(output_path, storage_path)

        html_url = ""
        html_name = ""
        try:
            html_output_name = safe_output.replace(".docx", ".html")
            html_output_path = tmp_dir_path / html_output_name
            _docx_to_html(output_path, html_output_path, inject_ids=True)
            html_url = _upload_html_to_storage(html_output_path, html_output_name)
            html_name = html_output_name
        except Exception as e:
            logger.warning("DOCX HTML 변환 실패: %s", e)

    if not file_url:
        raise RuntimeError("DOCX 업로드 실패")

    logger.info("generate_docx done: output=%s url=%s html_url=%s", safe_output, file_url, html_url)
    return {
        "file_name": safe_output,
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "file_url": file_url,
        "html_url": html_url,
        "html_name": html_name,
        "report_id": rid,
    }


@mcp.tool
async def edit_docx_page_html(
    docx_url: Annotated[str, Field(description="원본 DOCX URL")],
    page_number: Annotated[int, Field(description="수정할 물리적 페이지 번호 (1부터 시작). 문서 섹션 번호(예: 4.1)가 아닌 실제 페이지 순서. 모르면 1부터 시도")],
    instruction: Annotated[str, Field(description="수정 지시사항")],
    include_original: Annotated[Optional[bool], Field(description="응답에 원본 페이지 HTML 포함 여부")] = False,
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (자동 주입)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
) -> dict:
    """DOCX의 지정한 페이지를 지시사항대로 수정해 edits를 반환한다.

    주의: page_number는 물리적 페이지 번호(1, 2, 3...)이며, 문서 내 섹션/항목 번호(예: 4.1, 3.2)와 다르다.
    사용자가 '4.1절', '3번 항목' 등을 언급하면 해당 내용이 몇 페이지에 있는지 파악하여 올바른 page_number를 사용해야 한다.

    요구 사항:
    - page_number와 instruction이 모두 필요
    - 지시사항에는 '어떤 내용을 어떻게 수정할지'를 구체적으로 포함

    입력 예시:
    - page_number: 1
    - instruction: "표에서 '담당부서' 값을 'AI전략팀'으로 변경"
    """
    if not docx_url:
        raise ValueError("docx_url is required")
    if not page_number or page_number < 1:
        raise ValueError("page_number must be >= 1")
    if not instruction:
        raise ValueError("instruction is required")

    template_name = Path(urlparse(docx_url).path).name or "template.docx"
    logger.info("edit_docx_page_html start: template=%s page=%d", template_name, page_number)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        html_output_path = tmp_dir_path / f"page-edit-{template_name}.html"

        response = requests.get(docx_url, timeout=60)
        response.raise_for_status()
        template_path.write_bytes(response.content)

        _docx_to_html(template_path, html_output_path, inject_ids=True)
        html_text = html_output_path.read_text(encoding="utf-8")
        pages = extract_pages(html_text)
        if not pages:
            raise ValueError("페이지를 추출할 수 없습니다.")
        if page_number > len(pages):
            raise ValueError(
                f"page_number 범위 초과: 이 문서는 총 {len(pages)}페이지입니다. "
                f"page_number는 물리적 페이지 번호(1~{len(pages)})를 사용하세요. "
                f"문서 섹션 번호(예: 4.1)와 다릅니다."
            )
        original_page = pages[page_number - 1]

    prompt_sys, prompt_user = _build_page_edit_patch_prompt(original_page, instruction)
    edits_result = await asyncio.to_thread(_call_llm_json, prompt_sys, prompt_user, 0.2)
    if not isinstance(edits_result, dict):
        raise ValueError("LLM 결과가 올바르지 않습니다.")
    edits = edits_result.get("edits", [])
    if not isinstance(edits, list):
        edits = []

    _orig_fills, orig_ids = extract_fills_and_ids(original_page)
    orig_fills = _orig_fills or {}
    td_rows = _extract_td_rows(original_page)
    label_to_value_id: dict[str, int] = {}
    label_id_to_value_id: dict[int, int] = {}
    for row in td_rows:
        for idx, td_id in enumerate(row):
            next_idx = idx + 1
            if next_idx >= len(row):
                continue
            label_text = _normalize_label_text(str(orig_fills.get(td_id, "")))
            if not label_text:
                continue
            value_id = row[next_idx]
            label_to_value_id[label_text] = value_id
            label_id_to_value_id[td_id] = value_id
    normalized_edits = []
    for item in edits:
        if not isinstance(item, dict):
            continue
        raw_id = item.get("id")
        raw_label = item.get("label")
        new_text = item.get("new_text")
        if new_text is None:
            continue
        target_id: Optional[int] = None
        if raw_label:
            label_key = _normalize_label_text(str(raw_label))
            if label_key in label_to_value_id:
                target_id = label_to_value_id[label_key]
        if target_id is None and raw_id is not None:
            try:
                numeric_id = int(raw_id)
            except (TypeError, ValueError):
                continue
            if orig_ids and numeric_id not in orig_ids:
                continue
            target_id = label_id_to_value_id.get(numeric_id, numeric_id)
        if target_id is None:
            continue
        normalized_edits.append(
            {"id": target_id, "new_text": _normalize_patch_text(str(new_text))}
        )

    logger.info("edit_docx_page_html done: page=%d edits=%d", page_number, len(normalized_edits))
    payload = {
        "page_number": page_number,
        "edits": normalized_edits,
    }
    if include_original:
        payload["original_page_html"] = original_page
    return payload


@mcp.tool
async def save_docx_from_html(
    docx_url: Annotated[str, Field(description="원본 DOCX URL")],
    edited_html: Annotated[str, Field(description="편집된 HTML (data-id 포함)")],
    output_name: Annotated[Optional[str], Field(description="저장할 DOCX 파일명")] = "",
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (자동 주입)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
) -> dict:
    """편집된 HTML을 DOCX로 반영하고 스토리지 URL로 반환한다."""
    if not docx_url:
        raise ValueError("docx_url is required")
    if not edited_html:
        raise ValueError("edited_html is required")

    from datetime import datetime as _dt
    template_name = Path(urlparse(docx_url).path).name or "template.docx"
    safe = re.sub(r"[^0-9A-Za-z._-]+", "_", (output_name or template_name).replace(".docx", "")).strip("_") or "output"
    stamp = _dt.utcnow().strftime("%Y%m%d_%H%M%S")
    base_name = f"edited-{safe}_{stamp}"
    output_docx_name = f"{base_name}.docx"
    output_html_name = f"{base_name}.html"

    logger.info("save_docx_from_html start: template=%s output=%s html_len=%d", template_name, output_docx_name, len(edited_html))

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        output_path = tmp_dir_path / output_docx_name
        html_output_path = tmp_dir_path / output_html_name

        response = requests.get(docx_url, timeout=60)
        response.raise_for_status()
        template_path.write_bytes(response.content)

        await asyncio.to_thread(_apply_html_edits_to_docx, template_path, output_path, edited_html)

        storage_path = f"deep-research/{uuid.uuid4().hex}/{output_docx_name}"
        file_url = upload_docx_to_storage(output_path, storage_path)

        html_url = ""
        try:
            _docx_to_html(output_path, html_output_path, inject_ids=True)
            html_url = _upload_html_to_storage(html_output_path, output_html_name)
        except Exception as e:
            logger.warning("DOCX HTML 재변환 실패: %s", e)

    if not file_url:
        raise RuntimeError("DOCX 업로드 실패")

    logger.info("save_docx_from_html done: output=%s url=%s html_url=%s", output_docx_name, file_url, html_url)
    return {
        "file_name": output_docx_name,
        "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "file_url": file_url,
        "html_name": output_html_name if html_url else "",
        "html_content_type": HTML_CONTENT_TYPE if html_url else "",
        "html_url": html_url,
    }


# generate_slides 는 Gemini 이미지 생성에 의존하므로, 이미지 생성이 꺼져 있으면(폐쇄망 등)
# 도구 자체를 등록하지 않는다. 본문 LLM(LLM_PROVIDER)과는 무관하다.
_SLIDE_TOOL_ENABLED = bool(IMAGE_GENERATION_ENABLED)
if not _SLIDE_TOOL_ENABLED:
    logger.info(
        "generate_slides 도구 비활성화 (IMAGE_GENERATION_ENABLED=%s)",
        IMAGE_GENERATION_ENABLED,
    )

_slide_tool_decorator = mcp.tool if _SLIDE_TOOL_ENABLED else (lambda f: f)


@_slide_tool_decorator
async def generate_slides(
    report_markdown: Annotated[Optional[str], Field(description="보고서 마크다운 (이 값이 있으면 markdown 기반으로 슬라이드 생성)")] = "",
    hwpx_html_url: Annotated[Optional[str], Field(description="이전에 generate_hwpx로 생성한 결과의 html_url. 이 값이 있으면 해당 문서 내용을 읽어 슬라이드를 생성하며 report_markdown보다 우선 적용된다.")] = "",
    research_goal: Annotated[Optional[str], Field(description="리서치 목표 (report_markdown/hwpx_html_url 없을 때 사용)")] = "",
    outline_json: Annotated[Optional[str], Field(description="개요 목록 (JSON list[str], research_goal 모드에서 사용)")] = "",
    sources_json: Annotated[Optional[str], Field(description="소스 목록 (JSON list[dict], research_goal 모드에서 사용)")] = "",
    deck_title: Annotated[Optional[str], Field(description="덱 제목")] = "",
    slide_count: Annotated[Optional[int], Field(description="슬라이드 개수 (0이면 자동)")] = 0,
    style: Annotated[Optional[str], Field(description="스타일/색감 가이드")] = "",
    report_id: Annotated[Optional[str], Field(description="이미지 스토리지 경로용 ID")] = "",
    tenant_id: Annotated[Optional[str], Field(description="테넌트 ID (memento RAG 검색용)")] = "",
    user_jwt: Annotated[Optional[str], Field(description="사용자 JWT (자동 주입)")] = "",
    user_uid: Annotated[Optional[str], Field(description="사용자 UID (자동 주입)")] = "",
    user_email: Annotated[Optional[str], Field(description="사용자 이메일 (자동 주입)")] = "",
    room_id: Annotated[Optional[str], Field(description="채팅방 ID (자동 주입). 있으면 방에 업로드된 문서만 참조한다.")] = "",
    proc_inst_id: Annotated[Optional[str], Field(description="프로세스 인스턴스 ID. 있으면 해당 프로세스에 ingest된 문서만 참조한다.")] = "",
) -> dict:
    """슬라이드 마크다운과 이미지를 생성해 반환한다.

    generate_hwpx로 완성된 문서를 슬라이드로 변환하려면 해당 결과의 html_url을 hwpx_html_url에 전달한다.
    report_markdown을 직접 가진 경우에는 그것을 사용하고, 아무것도 없으면 research_goal로 새로 리서치한다.
    우선순위: hwpx_html_url > report_markdown > research_goal
    """
    import asyncio as _asyncio
    import json as _json
    import uuid as _uuid

    rid = report_id or _uuid.uuid4().hex
    sc = int(slide_count) if slide_count else None

    # hwpx_html_url이 있으면 HTML에서 텍스트를 추출해 report_markdown으로 사용
    if hwpx_html_url and hwpx_html_url.strip():
        try:
            from bs4 import BeautifulSoup
            resp = requests.get(hwpx_html_url.strip(), timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")
            # 불필요한 태그 제거
            for tag in soup(["script", "style", "head"]):
                tag.decompose()
            report_markdown = soup.get_text(separator="\n", strip=True)
            logger.info("generate_slides: hwpx_html_url에서 텍스트 추출 완료 (len=%d)", len(report_markdown))
        except Exception as exc:
            logger.warning("generate_slides: hwpx_html_url 텍스트 추출 실패: %s", exc)
            if not report_markdown:
                raise ValueError(f"hwpx_html_url 처리 실패: {exc}")

    if report_markdown:
        slide_md_raw = await _asyncio.to_thread(build_slide_markdown, report_markdown, sc, style or None)
    elif research_goal:
        outline = _json.loads(outline_json) if outline_json else []
        sources = _json.loads(sources_json) if sources_json else []

        # tenant_id가 있으면 memento에서 내부 지식 자료를 검색해 sources에 추가
        if (tenant_id or "").strip():
            try:
                from .memento import search_memento_smart
                logger.info("generate_slides: memento RAG 검색 시작 (tenant_id=%s)", tenant_id)
                memento_sources = await search_memento_smart(
                    query=research_goal,
                    outline=outline if outline else [research_goal],
                    tenant_id=tenant_id.strip(),
                    room_id=(room_id or "").strip() or None,
                    proc_inst_id=(proc_inst_id or "").strip() or None,
                )
                if memento_sources:
                    sources = memento_sources + sources
                    logger.info("generate_slides: memento 소스 %d개 추가 (총 %d개)", len(memento_sources), len(sources))
                else:
                    logger.info("generate_slides: memento 검색 결과 없음")
            except Exception as exc:
                logger.warning("generate_slides: memento 검색 실패 (Tavily 폴백): %s", exc)

        # outline/sources가 여전히 비어있으면 Tavily로 자동 웹 검색
        if not outline and not sources:
            from .search import research_for_slides
            logger.info("generate_slides: 웹 검색 시작 — %s", research_goal)
            _outline, _sources = await _asyncio.to_thread(research_for_slides, research_goal)
            outline = _outline or outline
            sources = _sources or sources
            logger.info("generate_slides: 웹 검색 완료 — outline=%d, sources=%d", len(outline), len(sources))

        slide_md_raw = await _asyncio.to_thread(
            build_slide_markdown_from_research,
            research_goal, outline, sources, deck_title or "", sc, style or None,
        )
    else:
        raise ValueError("report_markdown 또는 research_goal 중 하나는 필수입니다.")

    logger.info("generate_slides: slide_md_raw len=%d report_id=%s", len(slide_md_raw), rid)

    slides_for_style = parse_slides(slide_md_raw)
    style_guide = await _asyncio.to_thread(build_style_guide, slides_for_style, deck_title or "", style or "")
    slide_markdown, image_urls = await _asyncio.to_thread(
        generate_slide_images, slide_md_raw, rid, style_guide, deck_title or "", sc
    )

    logger.info("generate_slides done: slides=%d images=%d", len(slides_for_style), len(image_urls))
    return {
        "slide_markdown": slide_markdown,
        "image_urls": image_urls,
        "report_id": rid,
    }

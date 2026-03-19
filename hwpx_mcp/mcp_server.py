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

from .config import DEBUG_OUTPUT_DIR, DEBUG_OUTPUT_ENABLED, LOG_PATH
from .runner import process_hwpx_file
from .hwpx_to_html import hwpx_to_html
from .hwpx_edit import apply_html_edits_to_hwpx
from .core.html_pages import extract_pages, extract_first_page
from .core.html_edit import extract_fills_and_ids
from .agent.agent import _call_llm_text, _call_llm_json


def _setup_logging() -> logging.Logger:
    log_path = Path(LOG_PATH)
    if log_path.exists():
        log_path.unlink()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("process-gpt-office-mcp")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False

    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s process-gpt-office-mcp - %(message)s"
    )
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(fmt)
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(fmt)
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    return logger


logger = _setup_logging()
DEBUG_OUTPUT_PATH = Path(__file__).resolve().parent / DEBUG_OUTPUT_DIR
SUPABASE_URL = os.environ.get("SUPABASE_URL", "").strip()
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "").strip()
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




@mcp.tool
async def generate_hwpx(
    template_url: Annotated[str, Field(description="HWPX 템플릿 URL")],
    report_topic: Annotated[str, Field(description="보고서 주제")],
    report_description: Annotated[Optional[str], Field(description="보고서 상세 설명")] = "",
    reference_text: Annotated[Optional[str], Field(description="참고할 텍스트")] = "",
) -> dict:
    """HWPX 템플릿을 채워 스토리지 URL로 반환한다."""
    if not template_url:
        raise ValueError("template_url is required")
    if not report_topic:
        raise ValueError("report_topic is required")

    report_description = report_description or ""
    reference_text = reference_text or ""
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
        )

        if DEBUG_OUTPUT_ENABLED:
            DEBUG_OUTPUT_PATH.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(output_path, DEBUG_OUTPUT_PATH / output_name)
        file_url = _upload_hwpx_to_storage(output_path, output_name)
        html_url = ""
        try:
            hwpx_to_html(output_path, html_output_path, use_lineseg=False, inject_ids=True)
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
            raise ValueError(f"page_number 범위 초과: 최대 {len(pages)}")
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

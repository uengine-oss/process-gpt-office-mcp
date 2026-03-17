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

import requests
from fastmcp import FastMCP
from pydantic import Field
from supabase import create_client

from .config import DEBUG_OUTPUT_DIR, DEBUG_OUTPUT_ENABLED, LOG_PATH
from .runner import process_hwpx_file
from .hwpx_to_html import hwpx_to_html
from .hwpx_edit import apply_html_edits_to_hwpx


def _setup_logging() -> logging.Logger:
    log_path = Path(LOG_PATH)
    if log_path.exists():
        log_path.unlink()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("process-gpt-office-mcp")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

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

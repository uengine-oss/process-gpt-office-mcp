import base64
import logging
import shutil
import tempfile
from pathlib import Path
from typing import Annotated
from urllib.parse import urlparse

import requests
from fastmcp import FastMCP
from pydantic import Field

from .config import DEBUG_OUTPUT_DIR, DEBUG_OUTPUT_ENABLED, LOG_PATH
from .runner import process_hwpx_file


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

mcp = FastMCP("process-gpt-office-mcp")


def _safe_filename_from_url(url: str) -> str:
    name = Path(urlparse(url).path).name
    if not name:
        return "template.hwpx"
    return name


@mcp.tool
async def generate_hwpx(
    template_url: Annotated[str, Field(description="HWPX 템플릿 URL")],
    report_topic: Annotated[str, Field(description="보고서 주제")],
    report_description: Annotated[str, Field(description="보고서 상세 설명")] = "",
    reference_text: Annotated[str, Field(description="참고할 텍스트")] = "",
) -> dict:
    """HWPX 템플릿을 채워 base64로 반환한다."""
    if not template_url:
        raise ValueError("template_url is required")
    if not report_topic:
        raise ValueError("report_topic is required")

    template_name = _safe_filename_from_url(template_url)
    output_name = f"filled-{template_name}"
    logger.info("generate_hwpx start: template=%s", template_name)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        template_path = tmp_dir_path / template_name
        output_path = tmp_dir_path / output_name

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

        file_bytes = output_path.read_bytes()
        encoded = base64.b64encode(file_bytes).decode("ascii")

        if DEBUG_OUTPUT_ENABLED:
            DEBUG_OUTPUT_PATH.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(output_path, DEBUG_OUTPUT_PATH / output_name)

    logger.info("generate_hwpx done: output=%s size=%d", output_name, len(encoded))
    return {
        "file_name": output_name,
        "content_type": "application/vnd.hancom.hwpx",
        "base64_data": encoded,
    }

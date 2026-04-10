"""
hwpx-mcp 서버 설정.

.env에는 민감한 키 값만 둔다:
  SUPABASE_URL, SUPABASE_KEY, MEMENTO_SERVICE_URL, MEMENTO_DRIVE_FOLDER_ID
  OPENAI_API_KEY, OPENROUTER_API_KEY, GOOGLE_API_KEY

나머지 설정은 이 파일에서 직접 관리한다.
"""

import os
from pathlib import Path


# ─── .env 로드 (키 값 전용) ──────────────────────────────────────────

def _load_env_file() -> None:
    root = Path(__file__).resolve().parents[1]
    env_path = root / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#") or "=" not in raw:
            continue
        key, value = raw.split("=", 1)
        key = key.strip()
        value = value.strip().strip("'").strip('"')
        if key and key not in os.environ:
            os.environ[key] = value


_load_env_file()


def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# .env에서 읽는 값 (민감 정보 — 여기서는 변경하지 않음)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SUPABASE_URL: str = _env("SUPABASE_URL")
SUPABASE_KEY: str = _env("SUPABASE_KEY")
MEMENTO_SERVICE_URL: str = _env("MEMENTO_SERVICE_URL", "http://memento-service:8005")
MEMENTO_DRIVE_FOLDER_ID: str = _env("MEMENTO_DRIVE_FOLDER_ID")
OPENAI_API_KEY: str = _env("OPENAI_API_KEY")
OPENROUTER_API_KEY: str = _env("OPENROUTER_API_KEY")
GOOGLE_API_KEY: str = _env("GOOGLE_API_KEY")
TAVILY_API_KEY: str = _env("TAVILY_API_KEY")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# LLM 설정 (config에서 직접 관리)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 사용할 Provider: "openai" | "openrouter" | "gemini" | "custom"
LLM_PROVIDER: str = "openrouter"

# Provider별 모델명
OPENAI_MODEL_NAME: str = "gpt-5.1"
OPENROUTER_MODEL_NAME: str = "openai/gpt-oss-120b"
GEMINI_MODEL_NAME: str = "gemini-3.1-flash-image-preview"

# 폐쇄망/커스텀 설정 (custom provider 사용 시)
LLM_BASE_URL: str | None = None       # 예: "http://my-llm-server:8080/v1"
LLM_CUSTOM_MODEL_NAME: str = ""       # custom provider 모델명
LLM_API_KEY: str = _env("LLM_API_KEY")  # custom provider 전용 API key (.env에서 관리)

# 타임아웃
LLM_TIMEOUT_SECONDS: float = 600.0

# 최종 모델명 결정 (LLM_PROVIDER에 따라 자동 선택)
_MODEL_MAP = {
    "openai": OPENAI_MODEL_NAME,
    "openrouter": OPENROUTER_MODEL_NAME,
    "gemini": GEMINI_MODEL_NAME,
    "custom": LLM_CUSTOM_MODEL_NAME,
}
MODEL_NAME: str = _MODEL_MAP.get(LLM_PROVIDER, OPENAI_MODEL_NAME)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 기능 On/Off (config에서 직접 관리)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LLM_VISION_ENABLED: bool = False       # 비전 모델 사용 여부
WEB_SEARCH_ENABLED: bool = False       # Tavily 웹검색
IMAGE_GENERATION_ENABLED: bool = False  # Gemini 이미지 생성
IMAGE_REFERENCE_ENABLED: bool = False   # Memento 이미지 참조
VISION_CHUNK_PLAN_ENABLED: bool = False # 비전 기반 청크 분할 계획


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 이미지 생성 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GEMINI_IMAGE_MODEL: str = "gemini-3.1-flash-image-preview"
GEMINI_IMAGE_TIMEOUT_SECONDS: float = 120.0


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTTP 타임아웃 (곳곳에 하드코딩된 값을 여기서 관리)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HTTP_TIMEOUT_SHORT: float = 30.0   # 파일 다운로드, 검색 API 등
HTTP_TIMEOUT_LONG: float = 60.0    # 대용량 파일 다운로드, MCP 내부 호출

# OpenRouter reasoning effort: "low" | "medium" | "high" | None
REASONING_EFFORT: str | None = "low"

# Tavily API URL
TAVILY_API_URL: str = "https://api.tavily.com/search"

# 서버 포트
SERVER_PORT: int = 1192


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 레이아웃/파서 상수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
MAX_CONCURRENT_LLM: int = 10
SMALL_CELL_HEIGHT_MM: int = 3
SMALL_CELL_WIDTH_MM: int = 8
IMAGE_MIN_WIDTH_MM: int = 80
IMAGE_MIN_HEIGHT_MM: int = 25


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 디버그/로그
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LOG_PATH: str = "./run.log"
DEBUG_OUTPUT_ENABLED: bool = True
DEBUG_OUTPUT_DIR: str = "debug_outputs"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 로그 출력
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def log_config_summary() -> None:
    """서버 시작 시 주요 설정값을 로그로 출력한다."""
    import logging
    logger = logging.getLogger("process-gpt-office-mcp")
    lines = [
        "╔══════════════════════════════════════╗",
        "║        CONFIG SUMMARY                ║",
        "╠══════════════════════════════════════╣",
        f"  LLM_PROVIDER        = {LLM_PROVIDER}",
        f"  MODEL_NAME          = {MODEL_NAME}",
        f"  LLM_BASE_URL        = {LLM_BASE_URL or '(default)'}",
        f"  LLM_TIMEOUT_SECONDS = {LLM_TIMEOUT_SECONDS}",
        f"  LLM_VISION_ENABLED  = {LLM_VISION_ENABLED}",
        f"  WEB_SEARCH_ENABLED  = {WEB_SEARCH_ENABLED}",
        f"  IMAGE_GENERATION    = {IMAGE_GENERATION_ENABLED}",
        f"  IMAGE_REFERENCE     = {IMAGE_REFERENCE_ENABLED}",
        f"  VISION_CHUNK_PLAN   = {VISION_CHUNK_PLAN_ENABLED}",
        "╚══════════════════════════════════════╝",
    ]
    for line in lines:
        logger.info(line)

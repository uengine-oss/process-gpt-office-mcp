"""
hwpx-mcp 서버 설정.

설계 원칙:
  - .env: 키, 프로바이더 선택, URL (민감 정보 / 배포 환경별 엔드포인트)
  - 이 파일: 모델명, 기능 on/off, 타임아웃, 포트 등 앱 동작 상수

폐쇄망(air-gapped) 전환 시 `.env`에서 OFFICE_MCP_LLM_PROVIDER=custom 과
CUSTOM_LLM_BASE_URL / CUSTOM_LLM_API_KEY 만 바꾸면 된다.
"""

import os
from pathlib import Path


# ─── .env 로드 ──────────────────────────────────────────

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


def _env_optional(name: str) -> str | None:
    v = _env(name)
    return v or None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# .env 로부터 읽어오는 값 (키 / 프로바이더 / URL)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 인프라 URL
SUPABASE_URL: str = _env("SUPABASE_URL")
MEMENTO_SERVICE_URL: str = _env("MEMENTO_SERVICE_URL", "http://memento-service:8005")
TAVILY_API_URL: str = _env("TAVILY_API_URL", "https://api.tavily.com/search")

# API Keys
SUPABASE_KEY: str = _env("SUPABASE_KEY")
OPENAI_API_KEY: str = _env("OPENAI_API_KEY")
OPENROUTER_API_KEY: str = _env("OPENROUTER_API_KEY")
GOOGLE_API_KEY: str = _env("GOOGLE_API_KEY")
TAVILY_API_KEY: str = _env("TAVILY_API_KEY")

# LLM Provider 선택: openai | openrouter | gemini | custom
LLM_PROVIDER: str = _env("OFFICE_MCP_LLM_PROVIDER", "openrouter").lower()

# Custom provider (폐쇄망) URL + key
# 주의: LLM_PROVIDER=custom 일 때만 쓰인다. openrouter/openai 호출에 섞이면 안 된다.
if LLM_PROVIDER == "custom":
    LLM_BASE_URL: str | None = _env_optional("CUSTOM_LLM_BASE_URL")
    LLM_API_KEY: str = _env("CUSTOM_LLM_API_KEY")
else:
    LLM_BASE_URL = None
    LLM_API_KEY = ""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 앱 상수 (코드에서 관리 — 배포마다 바뀌지 않음)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Provider별 모델명 (LLM_PROVIDER 에 따라 자동 선택)
OPENAI_MODEL_NAME: str = "gpt-5.1"
OPENROUTER_MODEL_NAME: str = "openai/gpt-oss-120b"
GEMINI_MODEL_NAME: str = "gemini-3.1-flash-image-preview"
LLM_CUSTOM_MODEL_NAME: str = "/models/openai/gpt-oss-120b"

_MODEL_MAP = {
    "openai": OPENAI_MODEL_NAME,
    "openrouter": OPENROUTER_MODEL_NAME,
    "gemini": GEMINI_MODEL_NAME,
    "custom": LLM_CUSTOM_MODEL_NAME,
}
MODEL_NAME: str = _MODEL_MAP.get(LLM_PROVIDER, OPENAI_MODEL_NAME)

# LLM 공통
LLM_TIMEOUT_SECONDS: float = 600.0
REASONING_EFFORT: str | None = "low"  # "low" | "medium" | "high" | None

# 기능 On/Off
LLM_VISION_ENABLED: bool = False
WEB_SEARCH_ENABLED: bool = False
IMAGE_GENERATION_ENABLED: bool = False
IMAGE_REFERENCE_ENABLED: bool = False
VISION_CHUNK_PLAN_ENABLED: bool = False

# 이미지 생성 (Gemini)
GEMINI_IMAGE_MODEL: str = "gemini-3.1-flash-image-preview"
GEMINI_IMAGE_TIMEOUT_SECONDS: float = 120.0

# 서버
SERVER_PORT: int = 1192

# Memento
MEMENTO_DRIVE_FOLDER_ID: str = "1hUBPAOp-UVUicb-X_Cd69bbWGi4_vvM1"

# HTTP 타임아웃
HTTP_TIMEOUT_SHORT: float = 30.0   # 파일 다운로드, 검색 API 등
HTTP_TIMEOUT_LONG: float = 60.0    # 대용량 파일 다운로드, MCP 내부 호출

# LLM 동시 호출 제한
MAX_CONCURRENT_LLM: int = 10

# 레이아웃/파서 상수
SMALL_CELL_HEIGHT_MM: int = 3
SMALL_CELL_WIDTH_MM: int = 8
IMAGE_MIN_WIDTH_MM: int = 80
IMAGE_MIN_HEIGHT_MM: int = 25

# 디버그/로그 경로
LOG_PATH: str = "./run.log"
DEBUG_OUTPUT_ENABLED: bool = True
DEBUG_OUTPUT_DIR: str = "debug_outputs"


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 로그 출력
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _mask_secret(value: str) -> str:
    if not value:
        return "<empty>"
    if len(value) <= 8:
        return "***"
    return f"{value[:4]}...{value[-4:]} (len={len(value)})"


def log_config_summary() -> None:
    """서버 시작 시 주요 설정값을 로그로 출력한다."""
    import logging
    logger = logging.getLogger(__name__)
    lines = [
        "",
        "============================================================",
        "                    CONFIG SUMMARY",
        "============================================================",
        f"  LLM_PROVIDER            = {LLM_PROVIDER}",
        f"  MODEL_NAME              = {MODEL_NAME}",
        f"  LLM_BASE_URL            = {LLM_BASE_URL or '(provider default)'}",
        f"  LLM_API_KEY (custom)    = {_mask_secret(LLM_API_KEY)}",
        f"  LLM_TIMEOUT_SECONDS     = {LLM_TIMEOUT_SECONDS}",
        f"  REASONING_EFFORT        = {REASONING_EFFORT}",
        "  --- Feature flags ---",
        f"  LLM_VISION_ENABLED      = {LLM_VISION_ENABLED}",
        f"  WEB_SEARCH_ENABLED      = {WEB_SEARCH_ENABLED}",
        f"  IMAGE_GENERATION        = {IMAGE_GENERATION_ENABLED}",
        f"  IMAGE_REFERENCE         = {IMAGE_REFERENCE_ENABLED}",
        f"  VISION_CHUNK_PLAN       = {VISION_CHUNK_PLAN_ENABLED}",
        "  --- Infra ---",
        f"  MEMENTO_SERVICE_URL     = {MEMENTO_SERVICE_URL}",
        f"  SERVER_PORT             = {SERVER_PORT}",
        "============================================================",
    ]
    logger.info("\n".join(lines))

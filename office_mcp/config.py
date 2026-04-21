"""
hwpx-mcp 서버 설정.

설계 원칙:
  - 배포마다 바뀌는 값 (provider 선택, 모델명, base_url, 기능 on/off 등) → .env
  - 앱 동작 튜닝 상수 (레이아웃, HTTP 타임아웃, 경로 등) → 이 파일

폐쇄망(air-gapped) 전환 시 `.env`만 수정하고 컨테이너 재시작하면 됩니다.
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


def _env_bool(name: str, default: bool = False) -> bool:
    v = _env(name)
    if not v:
        return default
    return v.lower() in ("1", "true", "yes", "on", "y", "t")


def _env_float(name: str, default: float) -> float:
    v = _env(name)
    if not v:
        return default
    try:
        return float(v)
    except ValueError:
        return default


def _env_int(name: str, default: int) -> int:
    v = _env(name)
    if not v:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _env_optional(name: str) -> str | None:
    v = _env(name)
    return v or None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 인프라 / API 키 (민감 정보, .env)
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
# LLM 설정 (.env)
#   LLM_PROVIDER 로 선택: openai | openrouter | gemini | custom
#   provider별 모델명은 *_LLM_MODEL (네임스페이스 분리 — 모두 공존 가능)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LLM_PROVIDER: str = _env("LLM_PROVIDER", "openrouter").lower()

# Provider별 모델명 (기본값은 과거 config.py 값 유지)
OPENAI_MODEL_NAME: str = _env("OPENAI_LLM_MODEL", "gpt-5.1")
OPENROUTER_MODEL_NAME: str = _env("OPENROUTER_LLM_MODEL", "openai/gpt-oss-120b")
GEMINI_MODEL_NAME: str = _env("GEMINI_LLM_MODEL", "gemini-3.1-flash-image-preview")
LLM_CUSTOM_MODEL_NAME: str = _env("CUSTOM_LLM_MODEL")

# Custom provider용 URL / API key (폐쇄망).
# 주의: LLM_PROVIDER=custom 일 때만 채워 넣는다. 그 외 provider는 SDK 기본 URL을 써야 하므로
#       여기 값이 섞이면 openrouter/openai 호출이 엉뚱한 호스트로 간다.
if LLM_PROVIDER == "custom":
    LLM_BASE_URL: str | None = _env_optional("CUSTOM_LLM_BASE_URL") or _env_optional("LLM_BASE_URL")
    LLM_API_KEY: str = _env("CUSTOM_LLM_API_KEY") or _env("LLM_API_KEY")
else:
    LLM_BASE_URL = None
    LLM_API_KEY = ""

# 공통 타임아웃
LLM_TIMEOUT_SECONDS: float = _env_float("LLM_TIMEOUT_SECONDS", 600.0)

# OpenRouter reasoning effort: "low" | "medium" | "high" | None
REASONING_EFFORT: str | None = _env_optional("REASONING_EFFORT") or "low"

# 최종 모델명 결정 (LLM_PROVIDER에 따라 자동 선택)
_MODEL_MAP = {
    "openai": OPENAI_MODEL_NAME,
    "openrouter": OPENROUTER_MODEL_NAME,
    "gemini": GEMINI_MODEL_NAME,
    "custom": LLM_CUSTOM_MODEL_NAME,
}
MODEL_NAME: str = _MODEL_MAP.get(LLM_PROVIDER, OPENAI_MODEL_NAME)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 기능 On/Off (.env — 배포 환경별로 다를 수 있음)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LLM_VISION_ENABLED: bool = _env_bool("LLM_VISION_ENABLED", False)
WEB_SEARCH_ENABLED: bool = _env_bool("WEB_SEARCH_ENABLED", False)
IMAGE_GENERATION_ENABLED: bool = _env_bool("IMAGE_GENERATION_ENABLED", False)
IMAGE_REFERENCE_ENABLED: bool = _env_bool("IMAGE_REFERENCE_ENABLED", False)
VISION_CHUNK_PLAN_ENABLED: bool = _env_bool("VISION_CHUNK_PLAN_ENABLED", False)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 이미지 생성 설정 (.env)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GEMINI_IMAGE_MODEL: str = _env("GEMINI_IMAGE_MODEL", "gemini-3.1-flash-image-preview")
GEMINI_IMAGE_TIMEOUT_SECONDS: float = _env_float("GEMINI_IMAGE_TIMEOUT_SECONDS", 120.0)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 서버 / 외부 URL (.env)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SERVER_PORT: int = _env_int("SERVER_PORT", 1192)
TAVILY_API_URL: str = _env("TAVILY_API_URL", "https://api.tavily.com/search")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 앱 튜닝 상수 (코드에서 관리 — 배포마다 바뀌지 않음)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
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
    logger = logging.getLogger("process-gpt-office-mcp")
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

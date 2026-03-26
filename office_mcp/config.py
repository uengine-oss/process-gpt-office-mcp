import os
from pathlib import Path


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


def _get_bool_env(name: str, default: bool) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "y", "on")


def _get_float_env(name: str, default: float) -> float:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return float(raw)
    except ValueError:
        return default


OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "").strip()
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY", "").strip()

# LLM provider: "openai" or "gemini"
LLM_PROVIDER = os.environ.get("LLM_PROVIDER", "openai").strip().lower()

# Model names per provider
OPENAI_MODEL_NAME = os.environ.get("OPENAI_MODEL_NAME", "gpt-5.2").strip()
GEMINI_MODEL_NAME = os.environ.get("GEMINI_MODEL_NAME", "gemini-3-flash-preview").strip()

# 현재 사용할 모델명 (provider에 따라 결정)
MODEL_NAME = GEMINI_MODEL_NAME if LLM_PROVIDER == "gemini" else OPENAI_MODEL_NAME

GEMINI_IMAGE_MODEL = "gemini-3.1-flash-image-preview"
OPENAI_TIMEOUT_SECONDS = _get_float_env("OPENAI_TIMEOUT_SECONDS", 300.0)
GEMINI_IMAGE_TIMEOUT_SECONDS = _get_float_env("GEMINI_IMAGE_TIMEOUT_SECONDS", 120.0)

MAX_CONCURRENT_LLM = 10
SMALL_CELL_HEIGHT_MM = 3
SMALL_CELL_WIDTH_MM = 8
IMAGE_MIN_WIDTH_MM = 80
IMAGE_MIN_HEIGHT_MM = 25

LOG_PATH = "./run.log"
IMAGE_GENERATION_ENABLED = _get_bool_env("IMAGE_GENERATION_ENABLED", False)
IMAGE_REFERENCE_ENABLED = _get_bool_env("IMAGE_REFERENCE_ENABLED", True)
VISION_CHUNK_PLAN_ENABLED = _get_bool_env("VISION_CHUNK_PLAN_ENABLED", True)
DEBUG_OUTPUT_ENABLED = True
DEBUG_OUTPUT_DIR = "debug_outputs"

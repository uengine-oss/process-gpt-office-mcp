"""LLM Provider abstraction layer.

Supports swappable backends via LLM_PROVIDER env var:
  openai     - OpenAI direct
  openrouter - OpenRouter (OpenAI-compatible)
  gemini     - Google Gemini
  custom     - Any OpenAI-compatible endpoint (air-gapped / on-prem)
"""

from __future__ import annotations

import json
import logging
import re
import time
from abc import ABC, abstractmethod

logger = logging.getLogger("process-gpt-office-mcp")

_MAX_RETRIES = 3

from ..config import REASONING_EFFORT


# ── helpers ──

def _extract_json_from_text(text: str) -> str:
    """Gemini 응답에서 JSON 블록을 추출한다 (```json ... ``` 또는 전체 텍스트)."""
    m = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if m:
        return m.group(1).strip()
    text = text.strip()
    if text.startswith("{"):
        return text
    return text


def _find_balanced_json(text: str) -> dict | None:
    """본문 내에 포함된 첫 번째로 parse되는 균형 잡힌 {...} 블록을 찾아 반환."""
    for i, c in enumerate(text):
        if c != "{":
            continue
        depth = 0
        in_str = False
        esc = False
        for j in range(i, len(text)):
            ch = text[j]
            if in_str:
                if esc:
                    esc = False
                elif ch == "\\":
                    esc = True
                elif ch == '"':
                    in_str = False
                continue
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[i:j + 1])
                    except Exception:
                        break
    return None


def _recover_empty_key_wrapper(text: str) -> dict | None:
    """gpt-oss 계열이 내놓는 외곽 wrapper 형태를 복원.

    모델이 실제 JSON을 `{ "": "frag1", "": "frag2", ... }` 중복 빈 키로 감싸서
    반환하는 경우가 있다. 외곽 `{\\n{` 이중 중괄호를 풀고, 각 빈 키 value를
    이어붙인 뒤 JSON으로 재파싱한다.
    """
    t = re.sub(r"^\s*\{\s*\n\s*\{", "{", text)
    # 모델이 string 종료 쿼트 대신 `\"\"`를 찍는 경우 복구
    t = re.sub(r'\\"\\"\s*\n?\s*(\}\s*)$', r'"\1', t)
    try:
        pairs = json.JSONDecoder(strict=False, object_pairs_hook=list).decode(t)
    except Exception:
        return None
    if not isinstance(pairs, list) or not pairs:
        return None
    fragments = [v for k, v in pairs if k == "" and isinstance(v, str)]
    if len(fragments) < 2:
        return None
    joined = "".join(fragments)
    for extra in ("", "}", "\n}"):
        try:
            result = json.loads(joined + extra, strict=False)
        except Exception:
            continue
        if isinstance(result, dict):
            return result
    return None


def _robust_json_loads(text: str) -> dict:
    """LLM 응답에서 JSON을 최대한 회복시켜 파싱.

    일부 모델(특히 reasoning 스트리밍)이 중첩/escape된 JSON을 내뱉는 경우가 있어
    다단 fallback으로 복구한다.
    """
    raw = _extract_json_from_text(text)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    # 1) 원본에서 균형 잡힌 JSON 블록 탐색
    found = _find_balanced_json(raw)
    if found is not None:
        return found
    # 2) escape 되어 문자열로 감싸진 JSON인 경우 unescape 후 재탐색
    unescaped = raw.replace('\\"', '"').replace("\\n", "\n").replace("\\t", "\t")
    found = _find_balanced_json(unescaped)
    if found is not None:
        return found
    # 3) gpt-oss 계열의 빈 키 wrapper 패턴 복원
    found = _recover_empty_key_wrapper(raw)
    if found is not None:
        return found
    # 회복 실패 — 원 에러 재발생
    return json.loads(raw)


# ── Abstract base ──

class LLMProvider(ABC):
    """Three-method contract for LLM backends.

    `schema`가 주어지면 provider-native structured output(OpenAI json_schema /
    Gemini response_schema)을 사용해 모델이 해당 스키마만 생성하도록 강제한다.
    형식은 OpenAI 규격:
        {"name": "<id>", "schema": {<JSON Schema object>}}
    strict 모드는 provider가 지원하면 기본 활성화한다. schema=None이면 기존의
    자유형 JSON 출력으로 폴백한다.
    """

    @abstractmethod
    def call_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        ...

    @abstractmethod
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        ...

    @abstractmethod
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        ...


# ── OpenAI-compatible (OpenAI / OpenRouter / Custom) ──

class OpenAIProvider(LLMProvider):
    """Covers OpenAI, OpenRouter, and any OpenAI-compatible endpoint."""

    def __init__(
        self,
        api_key: str,
        model_name: str,
        timeout: float = 300.0,
        base_url: str | None = None,
        default_headers: dict | None = None,
        reasoning_effort: str | None = None,
    ):
        from openai import OpenAI

        kwargs: dict = {"api_key": api_key, "timeout": timeout}
        if base_url:
            kwargs["base_url"] = base_url
        if default_headers:
            kwargs["default_headers"] = default_headers
        self._client = OpenAI(**kwargs)
        self._model = model_name
        self._reasoning_effort = reasoning_effort  # "low", "medium", "high", or None
        # gpt-5*, o1*, o3* 계열은 max_tokens 미지원 → max_completion_tokens 사용
        m = (model_name or "").lower()
        self._max_tokens_key = (
            "max_completion_tokens"
            if m.startswith(("gpt-5", "o1", "o3"))
            else "max_tokens"
        )

    # -- json --
    def call_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        return self._do_json(
            messages=[
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": prompt_user},
            ],
            temperature=temperature,
            schema=schema,
        )

    # -- vision json --
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        user_content: list[dict] = [{"type": "text", "text": prompt_user}]
        for img in images_b64:
            user_content.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img}", "detail": "high"},
            })
        return self._do_json(
            messages=[
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": user_content},
            ],
            temperature=temperature,
            schema=schema,
        )

    # -- text --
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        kwargs: dict = {
            "model": self._model,
            "messages": [
                {"role": "system", "content": prompt_sys},
                {"role": "user", "content": prompt_user},
            ],
            "temperature": temperature,
            self._max_tokens_key: 16384,
        }
        if self._reasoning_effort:
            kwargs["extra_body"] = {
                "reasoning": {"enabled": True, "effort": self._reasoning_effort},
            }
        resp = self._client.chat.completions.create(**kwargs)
        return resp.choices[0].message.content or ""

    # -- internal --
    def _build_response_format(self, schema: dict | None) -> dict:
        """Structured Outputs (json_schema)를 우선 사용하고, 없으면 json_object로 폴백."""
        if schema and isinstance(schema.get("schema"), dict):
            return {
                "type": "json_schema",
                "json_schema": {
                    "name": schema.get("name") or "response",
                    "strict": True,
                    "schema": schema["schema"],
                },
            }
        return {"type": "json_object"}

    def _do_json(
        self,
        messages: list[dict],
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        started = time.perf_counter()
        last_exc = None
        response_format = self._build_response_format(schema)
        for attempt in range(1, _MAX_RETRIES + 1):
            try:
                kwargs: dict = {
                    "model": self._model,
                    "messages": messages,
                    "temperature": temperature,
                    "response_format": response_format,
                    self._max_tokens_key: 16384,
                }
                if self._reasoning_effort:
                    kwargs["extra_body"] = {
                        "reasoning": {"enabled": True, "effort": self._reasoning_effort},
                    }
                # 스트리밍으로 호출 — reasoning 토큰을 실시간 콘솔 출력
                kwargs["stream"] = True
                stream = self._client.chat.completions.create(**kwargs)
                content_parts: list[str] = []
                reasoning_parts: list[str] = []
                finish_reason = "stop"
                for chunk in stream:
                    delta = chunk.choices[0].delta if chunk.choices else None
                    if delta:
                        rc = getattr(delta, "reasoning_content", None) or getattr(delta, "reasoning", None)
                        if rc:
                            reasoning_parts.append(rc)
                        if delta.content:
                            content_parts.append(delta.content)
                    if chunk.choices and chunk.choices[0].finish_reason:
                        finish_reason = chunk.choices[0].finish_reason

                elapsed = time.perf_counter() - started
                content = "".join(content_parts) or "{}"
                reasoning_text = "".join(reasoning_parts)
                if reasoning_text:
                    logger.info("[LLM REASONING] len=%d\n%s", len(reasoning_text), reasoning_text)
                logger.info(
                    "[LLM RAW] attempt=%d finish_reason=%s content_len=%d\n%s",
                    attempt, finish_reason, len(content), content,
                )
                # finish_reason이 length면 응답이 잘렸을 가능성
                if finish_reason == "length":
                    logger.warning(
                        "[LLM TRUNCATED] 응답이 max_tokens에 의해 잘림 (content_len=%d). 재시도.",
                        len(content),
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                # 모델이 ```json ... ``` 으로 감싸거나 중첩 escape된 경우 처리
                data = _robust_json_loads(content)
                # 빈 JSON 감지: _elapsed_s 외에 유의미한 키가 없으면 재시도
                meaningful_keys = [k for k in data if k != "_elapsed_s"]
                if not meaningful_keys:
                    logger.warning(
                        "[LLM EMPTY] 빈 JSON 응답 감지 (attempt=%d, raw=%s). 재시도.",
                        attempt, content[:500],
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                data["_elapsed_s"] = round(elapsed, 3)
                return data
            except Exception as exc:
                last_exc = exc
                raw_preview = content[:500] if "content" in dir() else "N/A"
                logger.warning(
                    "[LLM RETRY] attempt=%d/%d error=%s raw=%s",
                    attempt, _MAX_RETRIES, exc, raw_preview,
                )
                if attempt < _MAX_RETRIES:
                    time.sleep(1)
        # 모든 재시도 실패 — 마지막 예외가 있으면 raise, 아니면 빈 dict 반환
        if last_exc:
            raise last_exc
        return {"_elapsed_s": round(time.perf_counter() - started, 3)}


# ── Gemini ──

class GeminiProvider(LLMProvider):
    """Google Gemini via google-genai SDK."""

    def __init__(self, api_key: str, model_name: str):
        from google import genai
        from google.genai import types as genai_types

        self._client = genai.Client(api_key=api_key)
        self._model = model_name
        self._types = genai_types

    # -- json --
    def call_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        return self._do_json(prompt_sys, prompt_user, temperature, schema=schema)

    # -- vision json --
    def call_vision_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        images_b64: list[str],
        temperature: float,
        schema: dict | None = None,
    ) -> dict:
        return self._do_json(prompt_sys, prompt_user, temperature, images_b64=images_b64, schema=schema)

    # -- text --
    def call_text(self, prompt_sys: str, prompt_user: str, temperature: float) -> str:
        config = self._types.GenerateContentConfig(
            system_instruction=prompt_sys,
            temperature=temperature,
        )
        resp = self._client.models.generate_content(
            model=self._model,
            contents=prompt_user,
            config=config,
        )
        return resp.text or ""

    # -- internal --
    def _do_json(
        self,
        prompt_sys: str,
        prompt_user: str,
        temperature: float,
        images_b64: list[str] | None = None,
        schema: dict | None = None,
    ) -> dict:
        import base64

        contents: list = []
        contents.append(prompt_user)
        if images_b64:
            for img in images_b64:
                img_bytes = base64.b64decode(img)
                contents.append(
                    self._types.Part.from_bytes(data=img_bytes, mime_type="image/png")
                )

        config_kwargs: dict = {
            "system_instruction": prompt_sys,
            "temperature": temperature,
            "response_mime_type": "application/json",
        }
        if schema and isinstance(schema.get("schema"), dict):
            config_kwargs["response_schema"] = schema["schema"]
        config = self._types.GenerateContentConfig(**config_kwargs)

        started = time.perf_counter()
        last_exc = None
        for attempt in range(1, _MAX_RETRIES + 1):
            try:
                resp = self._client.models.generate_content(
                    model=self._model,
                    contents=contents,
                    config=config,
                )
                elapsed = time.perf_counter() - started
                raw_text = resp.text or "{}"
                logger.info(
                    "[LLM RAW] attempt=%d content_len=%d content_preview=%s",
                    attempt, len(raw_text), raw_text[:500],
                )
                data = _robust_json_loads(raw_text)
                # 빈 JSON 감지: _elapsed_s 외에 유의미한 키가 없으면 재시도
                meaningful_keys = [k for k in data if k != "_elapsed_s"]
                if not meaningful_keys:
                    logger.warning(
                        "[LLM EMPTY] 빈 JSON 응답 감지 (attempt=%d, raw=%s). 재시도.",
                        attempt, raw_text[:500],
                    )
                    if attempt < _MAX_RETRIES:
                        time.sleep(1)
                        continue
                data["_elapsed_s"] = round(elapsed, 3)
                return data
            except Exception as exc:
                last_exc = exc
                raw_preview = raw_text[:500] if "raw_text" in dir() else "N/A"
                logger.warning(
                    "[LLM RETRY] attempt=%d/%d error=%s raw=%s",
                    attempt, _MAX_RETRIES, exc, raw_preview,
                )
                if attempt < _MAX_RETRIES:
                    time.sleep(1)
        if last_exc:
            raise last_exc
        return {"_elapsed_s": round(time.perf_counter() - started, 3)}


# ── Factory ──

def create_provider() -> LLMProvider:
    """LLM_PROVIDER 설정에 따라 적절한 provider를 생성한다.

    키 우선순위: LLM_API_KEY > provider별 키 (OPENAI_API_KEY 등)
    URL 우선순위: LLM_BASE_URL > provider별 기본값
    """
    from ..config import (
        LLM_PROVIDER, MODEL_NAME,
        LLM_API_KEY, LLM_BASE_URL, LLM_TIMEOUT_SECONDS,
        OPENAI_API_KEY,
        OPENROUTER_API_KEY,
        GOOGLE_API_KEY,
    )

    if LLM_PROVIDER == "gemini":
        api_key = GOOGLE_API_KEY
        if not api_key:
            raise RuntimeError("LLM_PROVIDER=gemini이지만 API 키가 없습니다 (LLM_API_KEY 또는 GOOGLE_API_KEY)")
        logger.info("[LLM] Provider=gemini, model=%s", MODEL_NAME)
        return GeminiProvider(api_key=api_key, model_name=MODEL_NAME)

    if LLM_PROVIDER == "openrouter":
        api_key = OPENROUTER_API_KEY
        if not api_key:
            raise RuntimeError("LLM_PROVIDER=openrouter이지만 API 키가 없습니다 (LLM_API_KEY 또는 OPENROUTER_API_KEY)")
        base_url = LLM_BASE_URL or "https://openrouter.ai/api/v1"
        logger.info("[LLM] Provider=openrouter, model=%s, reasoning_effort=%s", MODEL_NAME, REASONING_EFFORT)
        return OpenAIProvider(
            api_key=api_key,
            base_url=base_url,
            timeout=LLM_TIMEOUT_SECONDS,
            model_name=MODEL_NAME,
            reasoning_effort=REASONING_EFFORT,
        )

    if LLM_PROVIDER == "custom":
        base_url = LLM_BASE_URL
        if not base_url:
            raise RuntimeError("LLM_PROVIDER=custom이지만 LLM_BASE_URL이 없습니다")
        api_key = LLM_API_KEY or OPENAI_API_KEY or "not-needed"
        logger.info("[LLM] Provider=custom, base_url=%s, model=%s", base_url, MODEL_NAME)
        return OpenAIProvider(
            api_key=api_key,
            base_url=base_url,
            timeout=LLM_TIMEOUT_SECONDS,
            model_name=MODEL_NAME,
        )

    # default: openai
    api_key = OPENAI_API_KEY
    if not api_key:
        raise RuntimeError("LLM_PROVIDER=openai이지만 API 키가 없습니다 (LLM_API_KEY 또는 OPENAI_API_KEY)")
    logger.info("[LLM] Provider=openai, model=%s", MODEL_NAME)
    return OpenAIProvider(
        api_key=api_key,
        base_url=LLM_BASE_URL,
        timeout=LLM_TIMEOUT_SECONDS,
        model_name=MODEL_NAME,
    )

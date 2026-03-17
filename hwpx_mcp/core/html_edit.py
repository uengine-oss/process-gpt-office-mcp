from html import unescape
from html.parser import HTMLParser
import re


class _DataIdHtmlParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._id_stack: list[str | None] = []
        self._text_map: dict[str, list[str]] = {}
        self._seen_ids: set[str] = set()

    @property
    def text_map(self) -> dict[str, str]:
        return {k: "".join(v) for k, v in self._text_map.items()}

    @property
    def seen_ids(self) -> set[str]:
        return set(self._seen_ids)

    def _current_id(self) -> str | None:
        if not self._id_stack:
            return None
        return self._id_stack[-1]

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_map = {k: v for k, v in attrs}
        data_id = attrs_map.get("data-id")
        if data_id is not None:
            self._seen_ids.add(str(data_id))
        if data_id is None:
            data_id = self._current_id()
        self._id_stack.append(data_id)
        if tag.lower() == "br" and data_id is not None:
            self._text_map.setdefault(data_id, []).append("\n")

    def handle_endtag(self, tag: str) -> None:
        if self._id_stack:
            self._id_stack.pop()

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_map = {k: v for k, v in attrs}
        data_id = attrs_map.get("data-id") or self._current_id()
        if attrs_map.get("data-id") is not None:
            self._seen_ids.add(str(attrs_map.get("data-id")))
        if tag.lower() == "br" and data_id is not None:
            self._text_map.setdefault(data_id, []).append("\n")

    def handle_data(self, data: str) -> None:
        data_id = self._current_id()
        if data_id is None:
            return
        self._text_map.setdefault(data_id, []).append(data)


def _normalize_text(text: str) -> str:
    if not text:
        return ""
    value = unescape(text)
    value = value.replace("\u00a0", " ")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"\n{2,}", "\n", value)
    return value.strip()


def extract_fills_and_ids(edited_html: str) -> tuple[dict[int, str], set[int]]:
    if not edited_html:
        return {}, set()
    parser = _DataIdHtmlParser()
    parser.feed(edited_html)
    parser.close()
    output: dict[int, str] = {}
    for raw_id, text in parser.text_map.items():
        if raw_id is None:
            continue
        try:
            numeric = int(str(raw_id).strip())
        except ValueError:
            continue
        normalized = _normalize_text(text)
        output[numeric] = normalized
    present_ids: set[int] = set()
    for raw_id in parser.seen_ids:
        try:
            numeric = int(str(raw_id).strip())
        except ValueError:
            continue
        present_ids.add(numeric)
    return output, present_ids


def extract_fills_from_html(edited_html: str) -> dict[int, str]:
    fills, _present = extract_fills_and_ids(edited_html)
    return fills

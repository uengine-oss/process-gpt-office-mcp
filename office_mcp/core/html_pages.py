from html import escape
from html.parser import HTMLParser


def _attrs_to_text(attrs: list[tuple[str, str | None]]) -> str:
    parts: list[str] = []
    for key, value in attrs:
        if value is None:
            parts.append(key)
        else:
            parts.append(f'{key}="{escape(value, quote=True)}"')
    return " ".join(parts)


class _PageExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.pages: list[str] = []
        self._capturing = False
        self._depth = 0
        self._buffer: list[str] = []

    def _start_capture(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._capturing = True
        self._depth = 1
        attrs_text = _attrs_to_text(attrs)
        if attrs_text:
            self._buffer.append(f"<{tag} {attrs_text}>")
        else:
            self._buffer.append(f"<{tag}>")

    def _finish_capture(self, tag: str) -> None:
        self._buffer.append(f"</{tag}>")
        self.pages.append("".join(self._buffer))
        self._buffer = []
        self._capturing = False
        self._depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if not self._capturing:
            if tag.lower() != "div":
                return
            attrs_map = {k: v or "" for k, v in attrs}
            class_val = attrs_map.get("class", "")
            if "page" in class_val.split():
                self._start_capture(tag, attrs)
            return
        self._depth += 1
        attrs_text = _attrs_to_text(attrs)
        if attrs_text:
            self._buffer.append(f"<{tag} {attrs_text}>")
        else:
            self._buffer.append(f"<{tag}>")

    def handle_endtag(self, tag: str) -> None:
        if not self._capturing:
            return
        self._depth -= 1
        if self._depth <= 0:
            self._finish_capture(tag)
            return
        self._buffer.append(f"</{tag}>")

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if not self._capturing:
            return
        attrs_text = _attrs_to_text(attrs)
        if attrs_text:
            self._buffer.append(f"<{tag} {attrs_text} />")
        else:
            self._buffer.append(f"<{tag} />")

    def handle_data(self, data: str) -> None:
        if self._capturing:
            self._buffer.append(escape(data, quote=False))

    def handle_entityref(self, name: str) -> None:
        if self._capturing:
            self._buffer.append(f"&{name};")

    def handle_charref(self, name: str) -> None:
        if self._capturing:
            self._buffer.append(f"&#{name};")


def extract_pages(html_text: str) -> list[str]:
    if not html_text:
        return []
    parser = _PageExtractor()
    parser.feed(html_text)
    parser.close()
    return parser.pages


def extract_first_page(html_text: str) -> str:
    pages = extract_pages(html_text)
    return pages[0] if pages else ""

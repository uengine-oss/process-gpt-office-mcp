from __future__ import annotations

from dataclasses import dataclass


_HWP_UNITS_PER_MM = 283.465


@dataclass
class TableSummary:
    table_idx: int
    row_cnt: int
    col_cnt: int
    header_texts: list[str]
    data_min_width_mm: float
    empty_cell_ratio: float
    bfid_alternating: bool

    def summary_text(self) -> str:
        lines = [
            f"표{self.table_idx}: {self.row_cnt}행×{self.col_cnt}열",
            f"  헤더: {' | '.join(self.header_texts[:15])}",
            f"  데이터셀 최소너비: {self.data_min_width_mm:.1f}mm",
            f"  빈셀 비율: {self.empty_cell_ratio:.0%}",
            f"  bfid 교대패턴: {'있음' if self.bfid_alternating else '없음'}",
        ]
        return "\n".join(lines)


class TextNode:
    def __init__(
        self,
        nid: int,
        ntype: str,
        text: str,
        raw_text: str,
        depth: int,
        skip_fill: bool,
        t_elements: list,
        run_elements: list,
        table_idx: int = -1,
        row: int = -1,
        col: int = -1,
        cell_width: int = 0,
        cell_height: int = 0,
        cell_col_span: int = 1,
        cell_row_span: int = 1,
        style_refs: dict | None = None,
        style_info: dict | None = None,
        style_summary: str = "",
        style_missing: dict | None = None,
    ):
        self.id = nid
        self.type = ntype
        self.text = text
        self.raw_text = raw_text
        self.depth = depth
        self.skip_fill = skip_fill
        self.t_elements = t_elements
        self.run_elements = run_elements
        self.table_idx = table_idx
        self.row = row
        self.col = col
        self.cell_width = cell_width
        self.cell_height = cell_height
        self.cell_col_span = cell_col_span
        self.cell_row_span = cell_row_span
        self.style_refs = style_refs or {}
        self.style_info = style_info or {}
        self.style_summary = style_summary or ""
        self.style_missing = style_missing or {}

    @property
    def cell_width_mm(self) -> int:
        if self.cell_width <= 0:
            return 0
        return round(self.cell_width / _HWP_UNITS_PER_MM)

    @property
    def cell_height_mm(self) -> int:
        if self.cell_height <= 0:
            return 0
        return round(self.cell_height / _HWP_UNITS_PER_MM)

    def display(self) -> str:
        if self.type == "table_cell":
            w = self.cell_width_mm
            h = self.cell_height_mm
            size = ""
            if w > 0 and h > 0:
                size = f"~{w}x{h}mm"
            elif w > 0:
                size = f"~{w}mm"
            elif h > 0:
                size = f"~h{h}mm"
            loc = f"표{self.table_idx}[R{self.row},C{self.col}]{size}"
        else:
            loc = f"본문[L{self.depth}]" if self.depth > 0 else "본문"
        style = f" {self.style_summary}" if self.style_summary else ""
        txt = self.text if self.text else "<빈칸>"
        return f"[{self.id:3d}] ({loc:20s}){style} {txt}"

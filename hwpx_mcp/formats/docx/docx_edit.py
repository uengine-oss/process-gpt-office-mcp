"""Apply HTML edits back to a DOCX file.

Strategy
--------
We walk the DOCX body in the same order as docx_to_html (inject_ids=True) to
assign the same sequential IDs, then replace text content in elements whose
ID appears in the edited HTML.
"""
from __future__ import annotations

import logging
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict

from ...core.html_edit import extract_fills_and_ids

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

logger = logging.getLogger("process-gpt-office-mcp")


# ---------------------------------------------------------------------------
# Counter (mirrors docx_to_html._Counter)
# ---------------------------------------------------------------------------

class _Counter:
    def __init__(self):
        self.value = 0

    def next(self) -> int:
        v = self.value
        self.value += 1
        return v


# ---------------------------------------------------------------------------
# Apply edits to XML
# ---------------------------------------------------------------------------

def _set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    """Replace all runs in a paragraph with a single run containing text."""
    # Collect existing rPr from first run (to preserve formatting)
    first_rpr = None
    for r in paragraph.findall(f"{{{W}}}r"):
        rpr_el = r.find(f"{{{W}}}rPr")
        if rpr_el is not None:
            first_rpr = rpr_el
        paragraph.remove(r)

    # Create new run
    new_r = ET.SubElement(paragraph, f"{{{W}}}r")
    if first_rpr is not None:
        new_r.append(first_rpr)
    new_t = ET.SubElement(new_r, f"{{{W}}}t")
    new_t.text = text
    if text.startswith(" ") or text.endswith(" "):
        new_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _set_cell_text(tc: ET.Element, text: str) -> None:
    """Replace the first paragraph in a cell with plain text."""
    paragraphs = tc.findall(f"{{{W}}}p")
    if not paragraphs:
        return
    # Keep first paragraph, remove the rest
    for p in paragraphs[1:]:
        tc.remove(p)
    _set_paragraph_text(paragraphs[0], text)


def _apply_edits_to_body(body: ET.Element, edits: Dict[int, str]) -> None:
    counter = _Counter()

    for child in body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            eid = counter.next()
            if eid in edits:
                _set_paragraph_text(child, edits[eid])
        elif tag == "tbl":
            _apply_edits_to_table(child, counter, edits)


def _apply_edits_to_table(tbl: ET.Element, counter: _Counter, edits: Dict[int, str]) -> None:
    for tr in tbl.findall(f"{{{W}}}tr"):
        for tc in tr.findall(f"{{{W}}}tc"):
            tc_pr = tc.find(f"{{{W}}}tcPr")
            # Skip vMerge continuation cells (same as docx_to_html)
            if tc_pr is not None:
                v_merge = tc_pr.find(f"{{{W}}}vMerge")
                if v_merge is not None:
                    val = v_merge.attrib.get(f"{{{W}}}val", "continue")
                    if val == "continue":
                        continue

            eid = counter.next()
            if eid in edits:
                _set_cell_text(tc, edits[eid])
            # Always walk nested tables to keep counter in sync with docx_to_html
            for child in tc:
                ctag = child.tag.split("}")[-1]
                if ctag == "tbl":
                    _apply_edits_to_table(child, counter, edits)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def apply_html_edits_to_docx(
    docx_path: Path,
    output_path: Path,
    edited_html: str,
) -> None:
    """Read docx_path, apply edits from edited_html, write to output_path."""
    edits, present_ids = extract_fills_and_ids(edited_html)
    logger.info(
        "apply_html_edits_to_docx: html_len=%d present_ids=%d edits=%d",
        len(edited_html),
        len(present_ids),
        len(edits),
    )
    if not edits:
        logger.warning("apply_html_edits_to_docx: no data-id found in HTML, copying original")
        shutil.copy2(docx_path, output_path)
        return

    # Read all files from the zip
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        file_bytes: Dict[str, bytes] = {name: zin.read(name) for name in names}

    # Parse and patch document.xml
    doc_xml = file_bytes["word/document.xml"].decode("utf-8")
    # Register namespaces to avoid ns0: prefixes in output
    ET.register_namespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
    ET.register_namespace("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex")
    ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
    ET.register_namespace("aink", "http://schemas.microsoft.com/office/drawing/2016/ink")
    ET.register_namespace("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d")
    ET.register_namespace("o", "urn:schemas-microsoft-com:office:office")
    ET.register_namespace("oel", "http://schemas.microsoft.com/office/2019/extlst")
    ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ET.register_namespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math")
    ET.register_namespace("v", "urn:schemas-microsoft-com:vml")
    ET.register_namespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
    ET.register_namespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
    ET.register_namespace("w10", "urn:schemas-microsoft-com:office:word")
    ET.register_namespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    ET.register_namespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
    ET.register_namespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml")
    ET.register_namespace("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex")
    ET.register_namespace("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid")
    ET.register_namespace("w16", "http://schemas.microsoft.com/office/word/2018/wordml")
    ET.register_namespace("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash")
    ET.register_namespace("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex")
    ET.register_namespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
    ET.register_namespace("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk")
    ET.register_namespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml")
    ET.register_namespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")

    doc_root = ET.fromstring(doc_xml)
    body = doc_root.find(f"{{{W}}}body")
    if body is not None:
        _apply_edits_to_body(body, edits)

    patched_xml = ET.tostring(doc_root, encoding="unicode", xml_declaration=False)
    patched_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + patched_xml
    file_bytes["word/document.xml"] = patched_xml.encode("utf-8")

    # Write new zip
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, file_bytes[name])

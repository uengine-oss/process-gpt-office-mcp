import logging
import re
from xml.etree import ElementTree as ET

from ..models import TextNode
from ..images import apply_image_markers_to_section
from .parser import collect_runs_and_texts
from .xml_utils import find_parent, tag


logger = logging.getLogger("process-gpt-office-mcp")


def apply_fills(
    nodes: list[TextNode],
    fills: list[dict],
    tree: ET.ElementTree,
    section_path: str,
    parent_map: dict,
    instruction_ids: set[int],
    remove_table_indices: set[int],
    remove_node_ids: set[int] | None = None,
    t_ns: str = "",
    image_inserts: list[dict] | None = None,
    image_export_dir: str | None = None,
    export_prefix: str | None = None,
) -> None:
    fill_map = {
        f["id"]: f["new_text"]
        for f in fills
        if isinstance(f, dict) and "id" in f and "new_text" in f
    }
    node_ids_in_list = sorted([n.id for n in nodes if n.id is not None])
    missing_in_nodes = [k for k in fill_map if k not in node_ids_in_list]
    logger.info("[apply_fills] fill_map keys=%d, nodes=%d, fill_map에 있지만 nodes에 없는 ID=%s",
                len(fill_map), len(nodes), missing_in_nodes[:20] if missing_in_nodes else "없음")

    def _leading_ws(node: TextNode) -> str:
        if not node.raw_text:
            return ""
        raw = node.raw_text
        return raw[:len(raw) - len(raw.lstrip(" "))]

    def _normalize_checkbox(node: TextNode, new_text: str) -> str:
        return new_text

    def _table_first_cell_map() -> dict[int, str]:
        table_nodes: dict[int, list[TextNode]] = {}
        for n in nodes:
            if n.type == "table_cell":
                table_nodes.setdefault(n.table_idx, []).append(n)
        first_map: dict[int, str] = {}
        for tbl_idx, tbl_nodes in table_nodes.items():
            non_empty = [n for n in tbl_nodes if (n.text or "").strip()]
            if not non_empty:
                continue
            first_text = re.sub(r"\s+", " ", (non_empty[0].text or "").strip())
            if first_text:
                first_map[tbl_idx] = first_text
        return first_map

    table_first_cell = _table_first_cell_map()
    remove_table_signatures = {
        table_first_cell[tbl_idx]
        for tbl_idx in remove_table_indices
        if tbl_idx in table_first_cell and table_first_cell[tbl_idx]
    }

    def _remove_instruction_content(root: ET.Element):
        p_removed = 0
        tbl_removed = 0
        empty_removed = 0
        local_parent = {c: p for p in root.iter() for c in p}
        remove_ids = set(instruction_ids)
        if remove_node_ids:
            remove_ids.update(remove_node_ids)

        def _is_empty_p(p_elem: ET.Element) -> bool:
            for ch in p_elem.iter():
                if tag(ch) == "tbl":
                    return False
                if tag(ch) == "t" and (ch.text or "").strip():
                    return False
            return True

        def _safe_remove(parent_elem: ET.Element, child_elem: ET.Element) -> bool:
            if child_elem not in list(parent_elem):
                return False
            parent_elem.remove(child_elem)
            return True

        for node in nodes:
            if node.id not in remove_ids:
                continue
            if node.type != "body_text":
                continue
            target_run = None
            if node.t_elements:
                target_run = find_parent(node.t_elements[0], parent_map, "run")
            if target_run is None and node.run_elements:
                target_run = node.run_elements[0]
            if target_run is None:
                logger.warning("[정리] node %d: target_run=None (t=%d, run=%d) → 건너뜀",
                               node.id, len(node.t_elements), len(node.run_elements))
                continue
            parent_p = find_parent(target_run, parent_map, "p")
            if parent_p is None:
                continue
            parent = local_parent.get(parent_p)
            if parent is not None:
                siblings = list(parent)
                try:
                    idx = siblings.index(parent_p)
                except ValueError:
                    idx = -1
                if _safe_remove(parent, parent_p):
                    p_removed += 1
                if idx >= 0:
                    left_idx = idx - 1
                    while left_idx >= 0:
                        sib = siblings[left_idx]
                        if tag(sib) == "p" and _is_empty_p(sib):
                            if _safe_remove(parent, sib):
                                empty_removed += 1
                            left_idx -= 1
                            continue
                        break
                    right_idx = idx
                    while right_idx < len(siblings):
                        sib = siblings[right_idx]
                        if tag(sib) == "p" and _is_empty_p(sib):
                            if _safe_remove(parent, sib):
                                empty_removed += 1
                            right_idx += 1
                            continue
                        break

        tbl_index = -1
        logger.info("[정리] remove_table_indices=%s", remove_table_indices)
        for tbl in list(root.iter()):
            if tag(tbl) != "tbl":
                continue
            if find_parent(tbl, local_parent, "tbl") is not None:
                continue
            tbl_index += 1
            # 첫 셀 텍스트 (디버그용)
            _dbg_text = ""
            for _tc in tbl.iter():
                if tag(_tc) == "t" and (_tc.text or "").strip():
                    _dbg_text = (_tc.text or "").strip()[:20]
                    break
            if tbl_index in remove_table_indices:
                logger.info("[정리] 표%d 삭제 시도 (첫텍스트='%s')", tbl_index, _dbg_text)
                parent = local_parent.get(tbl)
                if parent is not None:
                    if _safe_remove(parent, tbl):
                        tbl_removed += 1
                        logger.info("[정리] 표%d 삭제 성공", tbl_index)
                    else:
                        logger.warning("[정리] 표%d 삭제 실패 (_safe_remove=False)", tbl_index)
                else:
                    logger.warning("[정리] 표%d 삭제 실패 (parent=None)", tbl_index)
                continue
            first_cell_text = ""
            for tc in tbl.iter():
                if tag(tc) != "tc":
                    continue
                _, t_elems = collect_runs_and_texts(tc)
                first_cell_text = "".join((t.text or "") for t in t_elems).strip()
                if first_cell_text:
                    break
            first_cell_norm = re.sub(r"\s+", " ", first_cell_text).strip()
            if first_cell_norm and first_cell_norm in remove_table_signatures:
                parent = local_parent.get(tbl)
                if parent is not None:
                    if _safe_remove(parent, tbl):
                        tbl_removed += 1

        if p_removed or tbl_removed:
            logger.info(
                "[정리] 삭제 완료: 문단 %d개, 표 %d개, 빈문단 %d개",
                p_removed,
                tbl_removed,
                empty_removed,
            )

    processed_p: set[ET.Element] = set()

    def _remove_linesegarray(p_elem: ET.Element | None) -> None:
        if p_elem is None or p_elem in processed_p:
            return
        for child in list(p_elem):
            if tag(child) == "linesegarray":
                p_elem.remove(child)
        processed_p.add(p_elem)

    for node in nodes:
        if node.id not in fill_map:
            continue

        new_text = _normalize_checkbox(node, fill_map[node.id])
        prefix = _leading_ws(node)
        if prefix:
            new_text = f"{prefix}{new_text}"

        # 원본 텍스트와 동일하면 XML 구조 보존을 위해 건너뜀 (공백 무시 비교)
        import re as _re
        if _re.sub(r"\s+", "", new_text) == _re.sub(r"\s+", "", node.raw_text or ""):
            logger.info("[fill] node %d: 원본과 동일 → 건너뜀", node.id)
            continue

        target_run = None
        if node.t_elements:
            target_run = find_parent(node.t_elements[0], parent_map, "run")
        if target_run is None and node.run_elements:
            target_run = node.run_elements[0]
        if target_run is None:
            logger.warning("[fill] node %d: target_run=None (t_elements=%d, run_elements=%d) → 건너뜀",
                           node.id, len(node.t_elements), len(node.run_elements))
            continue

        parent_p = find_parent(target_run, parent_map, "p")
        _remove_linesegarray(parent_p)

        if node.t_elements:
            node.t_elements[0].text = new_text
            logger.info("[fill] node %d: t_elements[0] 텍스트 교체 (%d자, t_elements=%d개)", node.id, len(new_text), len(node.t_elements))
            for extra_t in node.t_elements[1:]:
                extra_run = find_parent(extra_t, parent_map, "run")
                if extra_run is not None:
                    try:
                        extra_run.remove(extra_t)
                    except ValueError:
                        pass
                extra_t.text = ""
                extra_p = find_parent(extra_t, parent_map, "p")
                _remove_linesegarray(extra_p)
        elif node.run_elements:
            t_tag = f"{t_ns}t" if t_ns else "t"
            new_t = ET.Element(t_tag)
            target_run.insert(0, new_t)
            new_t.text = new_text
            logger.info("[fill] node %d: t 요소 새로 생성 (run에 삽입, %d자)", node.id, len(new_text))

    _remove_instruction_content(tree.getroot())

    if image_inserts:
        inserted = apply_image_markers_to_section(
            tree,
            section_path,
            parent_map,
            image_inserts,
            t_ns=t_ns,
            write_back=False,
            image_export_dir=image_export_dir,
            export_prefix=export_prefix,
        )
        if inserted:
            logger.info("[이미지] 마커 삽입 %d건", inserted)

    tree.write(section_path, encoding="utf-8", xml_declaration=True)

    with open(section_path, "r", encoding="utf-8") as f:
        raw = f.read()
    fixed = raw.replace(
        "<?xml version='1.0' encoding='utf-8'?>\r\n",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
        1,
    )
    if fixed == raw:
        fixed = raw.replace(
            "<?xml version='1.0' encoding='utf-8'?>\n",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
            1,
        )
    if fixed != raw:
        with open(section_path, "w", encoding="utf-8") as f:
            f.write(fixed)

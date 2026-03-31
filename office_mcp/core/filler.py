import logging
import os
import re
from xml.etree import ElementTree as ET

from ..models import TextNode
from ..images import apply_image_markers_to_section
from .parser import collect_runs_and_texts
from .xml_utils import find_parent, tag, register_namespaces


logger = logging.getLogger("process-gpt-office-mcp")

# ── 불릿 패턴: 내어쓰기(hanging indent)를 적용할 줄 패턴 ──
_BULLET_CHARS = set("ㅇ○●◦▪▸▹–—•·※☞◆◇■□▶►★☆")
_BULLET_RE = re.compile(
    r"^\s*[ㅇ○●◦▪▸▹\-–—•·※☞◆◇■□▶►★☆➊-➓❶-❿⓵-⓾①-⑳⑴-⒇]"
    r"|^\s*\d+[).]\s"
)

# 불릿 prefix 추출: leading whitespace + 불릿 기호 + 뒤따르는 공백
_BULLET_PREFIX_RE = re.compile(
    r"^(\s*(?:[ㅇ○●◦▪▸▹\-–—•·※☞◆◇■□▶►★☆➊-➓❶-❿⓵-⓾①-⑳⑴-⒇]|\d+[).])\s)"
)

# 한글 폰트에서 글자 폭 추정 계수.
# 관측값: height=1200 폰트에서 6글자 prefix의 intent=-9124 → 글자당 ~1520 → 계수 ~1.27
_CHAR_WIDTH_FACTOR = 1.27


def _extract_bullet_prefix(line: str) -> str | None:
    """줄에서 불릿 prefix를 추출한다.

    prefix = leading whitespace + 불릿 기호 + 뒤따르는 공백 1개.
    예: "    ㅇ 전기차..." → "    ㅇ "
        "  1) 첫번째..."  → "  1) "
    """
    m = _BULLET_PREFIX_RE.match(line)
    return m.group(1) if m else None


def _estimate_prefix_width(prefix: str, font_height: int) -> int:
    """prefix 문자열의 폭을 HWPUNIT으로 추정한다.

    한글 폰트 기준:
    - 모든 글자(공백 포함)를 font_height × _CHAR_WIDTH_FACTOR로 추정
    - 이 값은 한글 워드프로세서의 Shift+Tab 내어쓰기 결과와 근사함
    """
    return int(len(prefix) * font_height * _CHAR_WIDTH_FACTOR)


def _get_charpr_height(header_path: str, charpr_id: str) -> int | None:
    """header.xml에서 charPr의 font height를 읽는다."""
    if not header_path or not os.path.exists(header_path):
        return None
    tree = ET.parse(header_path)
    root = tree.getroot()
    for elem in root.iter():
        if tag(elem) == "charPr" and elem.get("id") == charpr_id:
            h = elem.get("height")
            return int(h) if h else None
    return None


def _find_or_create_bullet_parapr(
    header_path: str,
    base_parapr_id: str,
    indent_value: int,
) -> str | None:
    """header.xml에서 base_parapr_id와 동일하되 intent(내어쓰기)만 다른 paraPr를 찾거나 생성한다.

    Args:
        header_path: header.xml 경로
        base_parapr_id: 기준 paraPr의 id
        indent_value: 음수 HWPUNIT 값 (예: -9120)

    Returns:
        새/기존 내어쓰기용 paraPr의 id 문자열, 실패 시 None
    """
    if not header_path or not os.path.exists(header_path):
        return None

    register_namespaces(header_path)
    tree = ET.parse(header_path)
    root = tree.getroot()

    # ── 1) paraProperties 컨테이너와 기존 paraPr 목록 수집 ──
    para_props_container = None
    all_paraprs: list[ET.Element] = []
    base_elem: ET.Element | None = None

    for elem in root.iter():
        t = tag(elem)
        if t == "paraProperties":
            para_props_container = elem
        elif t == "paraPr":
            all_paraprs.append(elem)
            pid = elem.get("id")
            if pid == base_parapr_id:
                base_elem = elem

    if para_props_container is None or base_elem is None:
        return None

    # ── 2) base paraPr의 margin/intent 읽기 ──
    def _get_margin_child(parapr: ET.Element, child_tag: str) -> ET.Element | None:
        for ch in parapr:
            if tag(ch) == "margin":
                for sub in ch:
                    if tag(sub) == child_tag:
                        return sub
        return None

    base_intent = _get_margin_child(base_elem, "intent")
    base_intent_val = int(base_intent.get("value", "0")) if base_intent is not None else 0

    # 이미 내어쓰기가 적용된 paraPr이면 그대로 사용
    if base_intent_val != 0:
        return base_parapr_id

    # ── 3) 기존 paraPr 중 base와 동일하고 intent가 비슷한 것 찾기 ──
    def _margin_values(parapr: ET.Element) -> dict:
        vals = {}
        for ch in parapr:
            if tag(ch) == "margin":
                for sub in ch:
                    vals[tag(sub)] = sub.get("value", "0")
        return vals

    base_margins = _margin_values(base_elem)
    base_margins_no_intent = {k: v for k, v in base_margins.items() if k != "intent"}

    best_match: ET.Element | None = None
    best_diff = float("inf")

    for elem in all_paraprs:
        if elem is base_elem:
            continue
        m = _margin_values(elem)
        m_no_intent = {k: v for k, v in m.items() if k != "intent"}
        if m_no_intent != base_margins_no_intent:
            continue
        elem_intent = int(m.get("intent", "0"))
        if elem_intent == 0:
            continue
        # 다른 속성도 동일한지 비교 (align, lineSpacing 등)
        match = True
        for child_b in base_elem:
            ct = tag(child_b)
            if ct in ("align", "lineSpacing", "breakSetting", "autoSpacing"):
                found = False
                for child_c in elem:
                    if tag(child_c) == ct and child_c.attrib == child_b.attrib:
                        found = True
                        break
                if not found:
                    match = False
                    break
        if not match:
            continue

        # indent_value와 가장 가까운 paraPr 선택 (20% 오차 허용)
        diff = abs(elem_intent - indent_value)
        tolerance = abs(indent_value) * 0.20
        if diff <= tolerance and diff < best_diff:
            best_diff = diff
            best_match = elem

    if best_match is not None:
        reuse_id = best_match.get("id")
        reuse_intent = _margin_values(best_match).get("intent")
        logger.info("[bullet-indent] 기존 paraPr id=%s 재사용 (intent=%s, 요청=%s)",
                    reuse_id, reuse_intent, indent_value)
        return reuse_id

    # ── 4) 새 paraPr 생성 ──
    max_id = max((int(e.get("id", "0")) for e in all_paraprs), default=0)
    new_id = str(max_id + 1)

    import copy
    new_elem = copy.deepcopy(base_elem)
    new_elem.set("id", new_id)

    # margin 내 intent 값 설정
    intent_child = _get_margin_child(new_elem, "intent")
    if intent_child is not None:
        intent_child.set("value", str(indent_value))
    else:
        # margin 요소가 없으면 생성
        for ch in new_elem:
            if tag(ch) == "margin":
                intent_tag = ch[0].tag.rsplit("}", 1)[0] + "}intent" if "}" in ch[0].tag else "intent"
                new_intent = ET.SubElement(ch, intent_tag)
                new_intent.set("value", str(indent_value))
                new_intent.set("unit", "HWPUNIT")
                break

    para_props_container.append(new_elem)

    # itemCnt 업데이트
    cnt = para_props_container.get("itemCnt")
    if cnt is not None:
        para_props_container.set("itemCnt", str(int(cnt) + 1))

    tree.write(header_path, encoding="utf-8", xml_declaration=True)

    # XML declaration 스타일 맞추기
    with open(header_path, "r", encoding="utf-8") as f:
        raw = f.read()
    fixed = raw.replace(
        "<?xml version='1.0' encoding='utf-8'?>",
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
        1,
    )
    if fixed != raw:
        with open(header_path, "w", encoding="utf-8") as f:
            f.write(fixed)

    logger.info("[bullet-indent] 새 paraPr id=%s 생성 (base=%s, intent=%s)",
                new_id, base_parapr_id, indent_value)
    return new_id


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
    header_path: str | None = None,
    pregenerated_image_bytes: list[bytes | None] | None = None,
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

        # ── 본문 노드 + 여러 줄 → 줄마다 별도 <p> 문단 생성 ──
        lines = new_text.split("\n") if "\n" in new_text else []
        is_body_multiline = (
            node.type == "body_text"
            and len(lines) > 1
            and parent_p is not None
        )

        if is_body_multiline:
            # 원본 <p>의 속성 복사
            p_attribs = dict(parent_p.attrib)
            base_parapr_id = p_attribs.get("paraPrIDRef")
            char_pr = target_run.get("charPrIDRef", "")

            # ── font height 조회 (불릿 내어쓰기 폭 계산용) ──
            _font_height: int | None = None
            if header_path and char_pr:
                _font_height = _get_charpr_height(header_path, char_pr)

            # ── prefix별 bullet paraPr 캐시: {prefix_len: paraPrIDRef} ──
            _bullet_cache: dict[int, str | None] = {}

            def _get_bullet_parapr_for_line(line: str) -> str | None:
                """줄의 불릿 prefix를 분석하여 적절한 내어쓰기 paraPrIDRef를 반환."""
                if not header_path or not base_parapr_id:
                    return None
                prefix = _extract_bullet_prefix(line)
                if not prefix:
                    return None
                plen = len(prefix)
                if plen in _bullet_cache:
                    return _bullet_cache[plen]

                # prefix 폭 → intent 값 계산
                fh = _font_height or 1000  # fallback 10pt
                indent = -_estimate_prefix_width(prefix, fh)
                result = _find_or_create_bullet_parapr(
                    header_path, base_parapr_id, indent
                )
                _bullet_cache[plen] = result
                return result

            # 첫 줄은 원본 <p>에 작성
            first_line = lines[0]
            if node.t_elements:
                node.t_elements[0].text = first_line
                for extra_t in node.t_elements[1:]:
                    extra_t.text = ""
            elif node.run_elements:
                t_tag = f"{t_ns}t" if t_ns else "t"
                new_t = ET.Element(t_tag)
                target_run.insert(0, new_t)
                new_t.text = first_line

            # 첫 줄도 불릿이면 원본 <p>에 내어쓰기 적용
            bid = _get_bullet_parapr_for_line(first_line)
            if bid and bid != base_parapr_id:
                parent_p.set("paraPrIDRef", bid)

            # 나머지 줄은 새 <p> 문단으로 생성하여 원본 <p> 뒤에 삽입
            grandparent = find_parent(parent_p, parent_map, None)
            if grandparent is None:
                # parent_map에서 찾기
                for p, c_list in ((pp, list(pp)) for pp in tree.getroot().iter()):
                    if parent_p in c_list:
                        grandparent = p
                        break
            if grandparent is not None:
                insert_idx = list(grandparent).index(parent_p) + 1
                p_ns = parent_p.tag.rsplit("}", 1)[0] + "}" if "}" in parent_p.tag else ""
                run_ns = target_run.tag.rsplit("}", 1)[0] + "}" if "}" in target_run.tag else ""
                t_tag = f"{t_ns}t" if t_ns else "t"
                added = 0
                for line in lines[1:]:
                    # 불릿 줄이면 해당 prefix에 맞는 내어쓰기 paraPr 사용
                    bid = _get_bullet_parapr_for_line(line)
                    if bid:
                        line_attribs = dict(p_attribs)
                        line_attribs["paraPrIDRef"] = bid
                    else:
                        line_attribs = p_attribs

                    new_p = ET.Element(parent_p.tag, line_attribs)
                    new_run = ET.SubElement(new_p, target_run.tag)
                    if char_pr:
                        new_run.set("charPrIDRef", char_pr)
                    new_t_elem = ET.SubElement(new_run, t_tag)
                    new_t_elem.text = line
                    grandparent.insert(insert_idx, new_p)
                    insert_idx += 1
                    added += 1
                logger.info("[fill] node %d: 본문 %d줄 → %d개 <p> 문단 분리", node.id, len(lines), added + 1)
            else:
                # grandparent 못 찾으면 fallback: 한 덩어리로
                if node.t_elements:
                    node.t_elements[0].text = new_text
                logger.warning("[fill] node %d: grandparent 탐색 실패 → 단일 텍스트 fallback", node.id)
        elif node.type == "table_cell" and len(lines) > 1 and node.run_elements:
            # ── 표 셀 + 여러 줄 → 기존 문단 구조에 줄별 분배 ──
            # run_elements를 문서 순서대로 순회하며 소속 <hp:p>를 수집
            para_list: list[ET.Element] = []
            para_set: set[int] = set()
            para_t_map: dict[int, list[ET.Element]] = {}

            for run_el in node.run_elements:
                p_el = find_parent(run_el, parent_map, "p")
                if p_el is not None:
                    pid = id(p_el)
                    if pid not in para_set:
                        para_set.add(pid)
                        para_list.append(p_el)
                        para_t_map[pid] = []

            for t_el in node.t_elements:
                run_el = find_parent(t_el, parent_map, "run")
                if run_el:
                    p_el = find_parent(run_el, parent_map, "p")
                    if p_el is not None and id(p_el) in para_t_map:
                        para_t_map[id(p_el)].append(t_el)

            # 줄 수가 문단 수보다 많으면 마지막 문단에 나머지를 합침
            dist_lines = list(lines)
            if len(dist_lines) > len(para_list) and para_list:
                overflow_idx = len(para_list) - 1
                merged = "\n".join(dist_lines[overflow_idx:])
                dist_lines = dist_lines[:overflow_idx] + [merged]

            for i, line in enumerate(dist_lines):
                if i >= len(para_list):
                    break
                p_el = para_list[i]
                t_els = para_t_map.get(id(p_el), [])
                _remove_linesegarray(p_el)
                if t_els:
                    t_els[0].text = line
                    for extra_t in t_els[1:]:
                        extra_run = find_parent(extra_t, parent_map, "run")
                        if extra_run is not None:
                            try:
                                extra_run.remove(extra_t)
                            except ValueError:
                                pass
                        extra_t.text = ""
                else:
                    # 빈 문단 — 첫 번째 run에 <hp:t> 생성
                    for run_el in node.run_elements:
                        if find_parent(run_el, parent_map, "p") is p_el:
                            t_tag_str = f"{t_ns}t" if t_ns else "t"
                            new_t = ET.Element(t_tag_str)
                            new_t.text = line
                            run_el.insert(0, new_t)
                            break

            # 남은 문단(줄 수보다 많은 문단)은 텍스트 비우기
            for i in range(len(dist_lines), len(para_list)):
                p_el = para_list[i]
                t_els = para_t_map.get(id(p_el), [])
                _remove_linesegarray(p_el)
                for t_el in t_els:
                    t_el.text = ""

            logger.info("[fill] node %d: 표셀 %d줄 → %d개 문단에 분배",
                        node.id, len(dist_lines), min(len(dist_lines), len(para_list)))
        elif node.t_elements:
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
            pregenerated_bytes=pregenerated_image_bytes,
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

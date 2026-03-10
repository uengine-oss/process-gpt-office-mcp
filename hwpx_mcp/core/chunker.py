from ..models import TextNode


def chunk_nodes(nodes: list[TextNode], max_nodes: int) -> list[list[TextNode]]:
    if not nodes:
        return []

    table_groups: dict[int, list[TextNode]] = {}
    for n in nodes:
        if n.type == "table_cell":
            table_groups.setdefault(n.table_idx, []).append(n)

    placed_ids: set[int] = set()
    chunks: list[list[TextNode]] = []
    current: list[TextNode] = []

    for node in nodes:
        if node.id in placed_ids:
            continue

        if node.type == "table_cell":
            tbl_idx = node.table_idx
            table_cells = table_groups[tbl_idx]

            for tc in table_cells:
                placed_ids.add(tc.id)

            if len(table_cells) >= max_nodes:
                if current:
                    chunks.append(current)
                    current = []
                rows: dict[int, list[TextNode]] = {}
                for tc in table_cells:
                    rows.setdefault(tc.row, []).append(tc)
                row_group: list[TextNode] = []
                for row_cells in rows.values():
                    if row_group and len(row_group) + len(row_cells) > max_nodes:
                        chunks.append(row_group)
                        row_group = []
                    row_group.extend(row_cells)
                if row_group:
                    chunks.append(row_group)
                continue

            if current and len(current) + len(table_cells) > max_nodes:
                chunks.append(current)
                current = []

            current.extend(table_cells)

        else:
            placed_ids.add(node.id)
            if len(current) >= max_nodes:
                chunks.append(current)
                current = []
            current.append(node)

    if current:
        chunks.append(current)

    return chunks if chunks else [list(nodes)]


def chunk_nodes_by_plan(
    nodes: list[TextNode],
    plan: list[dict],
) -> tuple[list[list[TextNode]], list[str]]:
    errors: list[str] = []
    if not nodes:
        return [], ["empty nodes"]
    if not plan:
        return [], ["empty plan"]

    node_ids = [n.id for n in nodes if n.id is not None]
    node_id_set = set(node_ids)

    id_to_chunk: dict[int, int] = {}
    for idx, ch in enumerate(plan):
        ids = ch.get("node_ids") if isinstance(ch, dict) else None
        if not isinstance(ids, list) or not ids:
            errors.append(f"chunk {idx} has no node_ids")
            continue
        for nid in ids:
            if not isinstance(nid, int):
                errors.append(f"chunk {idx} has non-int id: {nid}")
                continue
            if nid in id_to_chunk:
                errors.append(f"duplicate id: {nid}")
                continue
            id_to_chunk[nid] = idx

    unknown_ids = [nid for nid in id_to_chunk.keys() if nid not in node_id_set]
    if unknown_ids:
        errors.append(f"unknown ids: {sorted(unknown_ids)[:20]}")

    missing_ids = [nid for nid in node_id_set if nid not in id_to_chunk]
    if missing_ids:
        errors.append(f"missing ids: {sorted(missing_ids)[:20]}")

    if errors:
        return [], errors

    chunks: list[list[TextNode]] = [[] for _ in range(len(plan))]
    for node in nodes:
        if node.id is None:
            continue
        idx = id_to_chunk.get(node.id)
        if idx is None:
            errors.append(f"unassigned id: {node.id}")
            continue
        chunks[idx].append(node)

    if any(len(c) == 0 for c in chunks):
        errors.append("empty chunk detected")

    if errors:
        return [], errors

    return chunks, []

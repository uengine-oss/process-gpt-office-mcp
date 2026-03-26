import logging
import re
from ..models import TextNode

logger = logging.getLogger("process-gpt-office-mcp")

# ── 청크 크기 제한 ──
_MAX_CHUNK_NODES = 120   # 이 이상이면 이어붙이지 않음
_MIN_CHUNK_NODES = 40    # 이 미만이면 인접 청크에 병합


def _is_heading(node: TextNode) -> bool:
    """본문 노드가 제목/섹션 구분자인지 판별."""
    if node.type != "body_text":
        return False
    text = (node.text or node.raw_text or "").strip()
    if not text:
        return False
    style = node.style_summary or ""
    # bold + 큰 폰트 = 제목
    if "bold" in style:
        m = re.search(r"size=(\d+)", style)
        if m and int(m.group(1)) >= 1200:
            return True
    # "Ⅰ.", "Ⅱ." 등 섹션 번호 패턴
    if text[0] in "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅰⅱⅲⅳ":
        return True
    return False


def _is_heading_table(table_cells: list[TextNode]) -> bool:
    """1~2행짜리 제목용 표인지 판별 (예: "Ⅰ. 개요", "1 추진배경")."""
    if not table_cells:
        return False
    rows = {n.row for n in table_cells}
    if len(rows) > 2:
        return False
    # 셀 수가 적고 (제목표는 보통 1~4셀)
    if len(table_cells) > 6:
        return False
    # 텍스트 중 제목 패턴이 있는지
    for n in table_cells:
        text = (n.text or n.raw_text or "").strip()
        if not text:
            continue
        if text[0] in "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅰⅱⅲⅳ":
            return True
        # "1", "2" 같은 숫자 + bold + 큰 폰트
        style = n.style_summary or ""
        if "bold" in style and text.isdigit():
            m = re.search(r"size=(\d+)", style)
            if m and int(m.group(1)) >= 1400:
                return True
    return False


def chunk_nodes_semantic(nodes: list[TextNode]) -> list[list[TextNode]]:
    """규칙 기반 시맨틱 청킹.

    규칙:
    1. 테이블은 절대 분리하지 않음 (같은 table_idx = 한 덩어리)
    2. 본문 제목(bold + 큰 폰트)을 만나면 새 청크 시작
    3. 테이블/텍스트를 이어붙일 때 MAX를 넘으면 이어붙이지 않음
    4. MIN 미만 청크는 인접 청크에 병합
    """
    if not nodes:
        return []

    # 테이블 그룹 미리 수집
    table_groups: dict[int, list[TextNode]] = {}
    for n in nodes:
        if n.type == "table_cell":
            table_groups.setdefault(n.table_idx, []).append(n)

    placed_ids: set[int] = set()
    chunks: list[list[TextNode]] = []
    current: list[TextNode] = []

    def _flush():
        nonlocal current
        if current:
            chunks.append(current)
            current = []

    i = 0
    while i < len(nodes):
        node = nodes[i]
        if node.id in placed_ids:
            i += 1
            continue

        # ── 테이블 노드: 테이블 전체를 한 번에 ──
        if node.type == "table_cell":
            tbl_idx = node.table_idx
            table_cells = table_groups.get(tbl_idx, [])
            for tc in table_cells:
                placed_ids.add(tc.id)

            # 제목용 표(1~2행)면 새 청크 시작
            if _is_heading_table(table_cells) and current:
                _flush()

            # 현재 청크 + 이 테이블이 MAX를 넘으면 flush
            # 단, 직전 텍스트 노드(제목/빈칸)는 테이블 쪽으로 가져감
            if current and len(current) + len(table_cells) > _MAX_CHUNK_NODES:
                # 현재 청크 끝에서 본문 노드를 역순으로 수집 (테이블 직전 텍스트)
                tail: list[TextNode] = []
                while current and getattr(current[-1], "type", "") == "body_text":
                    tail.append(current.pop())
                tail.reverse()
                _flush()
                current.extend(tail)  # 텍스트를 새 청크(테이블 앞)로 이동

            current.extend(table_cells)

            # 테이블 추가 후 MAX를 넘었으면 flush (다음 것과 붙지 않도록)
            if len(current) > _MAX_CHUNK_NODES:
                _flush()

            i += 1
            continue

        # ── 본문 노드 ──
        placed_ids.add(node.id)

        # 제목을 만나면 새 청크 시작 (현재 청크에 내용이 있을 때만)
        if _is_heading(node) and current:
            _flush()

        current.append(node)
        i += 1

    _flush()

    # ── 소청크 병합 ──
    if len(chunks) > 1:
        merged: list[list[TextNode]] = []
        for chunk in chunks:
            # 이전 청크가 MIN 미만이고, 합쳐도 MAX 이하면 병합
            if merged and len(merged[-1]) < _MIN_CHUNK_NODES:
                if len(merged[-1]) + len(chunk) <= _MAX_CHUNK_NODES:
                    merged[-1].extend(chunk)
                    continue
            merged.append(chunk)
        # 마지막이 MIN 미만이면 이전에 병합
        if len(merged) >= 2 and len(merged[-1]) < _MIN_CHUNK_NODES:
            if len(merged[-2]) + len(merged[-1]) <= _MAX_CHUNK_NODES:
                merged[-2].extend(merged.pop())
        if len(merged) != len(chunks):
            logger.info("[시맨틱청킹] %d → %d 청크 (소청크 병합)", len(chunks), len(merged))
        chunks = merged

    for idx, chunk in enumerate(chunks):
        table_count = len({n.table_idx for n in chunk if n.type == "table_cell" and n.table_idx >= 0})
        text_count = sum(1 for n in chunk if n.type == "body_text")
        logger.info("[시맨틱청킹] 청크 %d: %d노드 (표 %d개, 본문 %d개)", idx, len(chunk), table_count, text_count)

    logger.info("[시맨틱청킹] %d개 노드 → %d개 청크", len(nodes), len(chunks))
    return chunks


# ── 하위 호환용 ──

def chunk_nodes(nodes: list[TextNode], max_nodes: int) -> list[list[TextNode]]:
    return chunk_nodes_semantic(nodes)


def chunk_nodes_by_plan(
    nodes: list[TextNode],
    plan: list[dict],
) -> tuple[list[list[TextNode]], list[str]]:
    return [], ["deprecated: use chunk_nodes_semantic"]

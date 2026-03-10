import os
import re
import zipfile


_SIGNATURE_PATTERNS = ("sign", "signature", "encrypt", "certification", "xmlsig")


def extract_hwpx(hwpx_path: str, extract_dir: str) -> tuple[dict[str, int], list[str]]:
    with zipfile.ZipFile(hwpx_path, "r") as zf:
        infos = zf.infolist()
        compress_info = {info.filename: info.compress_type for info in infos}
        file_order = [info.filename for info in infos]
        zf.extractall(extract_dir)
    return compress_info, file_order


def _find_contents_dir(extract_dir: str) -> str | None:
    expected = os.path.join(extract_dir, "Contents")
    if os.path.isdir(expected):
        return expected
    try:
        for name in os.listdir(extract_dir):
            if name.lower() == "contents":
                candidate = os.path.join(extract_dir, name)
                if os.path.isdir(candidate):
                    return candidate
    except FileNotFoundError:
        return None
    return None


def find_section_files(extract_dir: str) -> list[str]:
    contents_dir = _find_contents_dir(extract_dir)
    sections: list[str] = []
    if contents_dir and os.path.isdir(contents_dir):
        sections = sorted([
            os.path.join(contents_dir, f)
            for f in os.listdir(contents_dir)
            if re.match(r"section.*\.xml", f, re.IGNORECASE)
        ])
    if sections:
        return sections

    # Fallback: search entire extracted tree for section*.xml
    matches: list[str] = []
    for root_dir, _dirs, files in os.walk(extract_dir):
        for f in files:
            if re.match(r"section.*\.xml", f, re.IGNORECASE):
                matches.append(os.path.join(root_dir, f))
    return sorted(matches)


def repack_hwpx(
    extract_dir: str,
    output_path: str,
    original_compress_info: dict[str, int] | None = None,
    original_file_order: list[str] | None = None,
) -> None:
    if os.path.exists(output_path):
        os.remove(output_path)

    all_files_map: dict[str, str] = {}
    for root_dir, _dirs, files in os.walk(extract_dir):
        for f in files:
            fp = os.path.join(root_dir, f)
            rel = os.path.relpath(fp, extract_dir).replace("\\", "/")
            all_files_map[rel] = fp

    def _compress_type(rel: str) -> int:
        if original_compress_info and rel in original_compress_info:
            return original_compress_info[rel]
        return zipfile.ZIP_DEFLATED

    def _should_skip(rel: str) -> bool:
        rel_lower = rel.lower()
        return any(p in rel_lower for p in _SIGNATURE_PATTERNS)

    if original_file_order:
        ordered_rels = list(original_file_order)
        original_set = set(original_file_order)
        for rel in all_files_map:
            if rel not in original_set:
                ordered_rels.append(rel)
    else:
        ordered_rels = list(all_files_map.keys())

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel in ordered_rels:
            if _should_skip(rel):
                continue
            fp = all_files_map.get(rel)
            if fp is None:
                continue
            zf.write(fp, rel, compress_type=_compress_type(rel))

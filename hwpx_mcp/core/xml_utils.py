from xml.etree import ElementTree as ET


def register_namespaces(xml_path: str) -> None:
    for _event, elem in ET.iterparse(xml_path, events=("start-ns",)):
        prefix, uri = elem
        if prefix is None:
            prefix = ""
        ET.register_namespace(prefix, uri)


def tag(elem: ET.Element) -> str:
    if "}" in elem.tag:
        return elem.tag.split("}", 1)[1]
    return elem.tag


def ns(elem: ET.Element) -> str:
    if elem.tag.startswith("{"):
        return elem.tag.split("}", 1)[0] + "}"
    return ""


def find_parent(node: ET.Element, parent_map: dict, tag_name: str) -> ET.Element | None:
    cur = node
    while cur is not None:
        if tag(cur) == tag_name:
            return cur
        cur = parent_map.get(cur)
    return None

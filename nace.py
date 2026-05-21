"""NACE Rev.2 sektör hiyerarşisi - naceler.txt'ten parse edilir."""
import re
from pathlib import Path

NACE_FILE = Path(__file__).parent / "naceler.txt"

RE_MAIN = re.compile(r"^\*\*([A-Z])\s+(.+?)\*\*$")
RE_TWO = re.compile(r"^([A-Z])\s+(\d{2})\s+(.+)$")
RE_THREE = re.compile(r"^([A-Z])\s+(\d{2})\.(\d)\s+(.+)$")


def load_nace_tree(path: Path = NACE_FILE):
    """
    Dön: { harf: {"name": str, "url_code": str, "children": {two_code: {...}}} }
    URL kodları: ana=harf, 2'li='10', 3'lü='101' (nokta yok).
    """
    tree = {}
    current_main = None

    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line:
            continue

        m = RE_MAIN.match(line)
        if m:
            letter, name = m.group(1), m.group(2).strip()
            current_main = letter
            tree[letter] = {"name": name, "url_code": letter, "children": {}}
            continue

        m = RE_THREE.match(line)
        if m and current_main == m.group(1):
            letter, two, three, name = m.group(1), m.group(2), m.group(3), m.group(4).strip()
            parent = tree[letter]["children"].get(two)
            if parent is None:
                continue
            kod_3 = f"{two}{three}"
            parent["children"][kod_3] = {
                "name": name,
                "url_code": kod_3,
                "display": f"{two}.{three}",
            }
            continue

        m = RE_TWO.match(line)
        if m and current_main == m.group(1):
            letter, two, name = m.group(1), m.group(2), m.group(3).strip()
            tree[letter]["children"][two] = {
                "name": name,
                "url_code": two,
                "display": two,
                "children": {},
            }
            continue

    return tree


def get_subsectors(tree, main_letter):
    """Bir ana sektörün altındaki 2'li ve 3'lü kodları liste olarak döner.
    Sıra: 2'li → onun 3'lüleri → sonraki 2'li → ...
    """
    if main_letter not in tree:
        return []
    options = []
    children = tree[main_letter]["children"]
    for two_code in sorted(children.keys()):
        two_data = children[two_code]
        options.append((two_data["url_code"], f"  {two_data['display']} - {two_data['name']}"))
        for three_code in sorted(two_data.get("children", {}).keys()):
            three_data = two_data["children"][three_code]
            options.append(
                (three_data["url_code"], f"      {three_data['display']} - {three_data['name']}")
            )
    return options

"""Parse ephemeral XML test outputs using a standards-compliant XML parser."""

import json
import sys
import xml.etree.ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"s": MAIN}


def inspect(document):
    root = ET.fromstring(document["xml"])
    kind = document["kind"]
    assert root.tag == f"{{{MAIN}}}{kind}", (document["label"], root.tag)
    children = [item.tag.removeprefix(f"{{{MAIN}}}") for item in root]
    result = {"label": document["label"], "children": children}
    if kind == "workbook":
        names = root.findall("s:definedNames", NS)
        assert len(names) == 1, (document["label"], "duplicate/missing definedNames")
        result["definedNames"] = [dict(item.attrib, text=item.text or "") for item in names[0]]
        assert children.index("sheets") < children.index("definedNames"), document["label"]
        if "calcPr" in children:
            assert children.index("definedNames") < children.index("calcPr"), document["label"]
    if kind == "worksheet":
        # Page elements must occur after sheet data and in SpreadsheetML order.
        positions = [children.index(name) for name in ("sheetData", "printOptions", "pageMargins", "pageSetup")]
        assert positions == sorted(positions), (document["label"], positions)
        for successor in ("headerFooter", "drawing", "legacyDrawing", "tableParts", "extLst"):
            if successor in children:
                assert children.index("pageSetup") < children.index(successor), document["label"]
        for name in ("printOptions", "pageMargins", "pageSetup"):
            elements = root.findall("s:" + name, NS)
            assert len(elements) == 1, (document["label"], name)
            result[name] = elements[0].attrib
        assert not any(key.endswith("}id") or key == "r:id" for key in result["pageSetup"]), (document["label"], "stale printer relationship")
        props = root.find("s:sheetPr/s:pageSetUpPr", NS)
        assert props is not None, (document["label"], "missing pageSetUpPr")
        result["pageSetUpPr"] = props.attrib
    if kind == "styleSheet":
        result["numFmts"] = [item.attrib for item in root.findall("s:numFmts/s:numFmt", NS)]
        fmt = root.find("s:numFmts", NS)
        assert fmt is not None
        assert int(fmt.attrib["count"]) == len(result["numFmts"]), document["label"]
    return result


print(json.dumps([inspect(item) for item in json.load(sys.stdin)["documents"]]))

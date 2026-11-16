"""Embed the authorized template and produce a deployment ZIP (stdlib only)."""
from pathlib import Path
import base64
import json
import re
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
APP = ROOT / "google_apps_script"
TEMPLATE = ROOT / "Tamplate_PL.xlsx"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def build_template():
    content = TEMPLATE.read_bytes()
    with zipfile.ZipFile(TEMPLATE) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        pl = next(sheet for sheet in workbook.find("s:sheets", NS) if sheet.get("name") == "PL")
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(rel.get("Target") for rel in relations if rel.get("Id") == pl.get(f"{{{R}}}id"))
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        sheet_xml = archive.read(path).decode("utf-8")
        settings = {}
        for tag in ("printOptions", "pageMargins", "pageSetup", "pageSetUpPr"):
            match = re.search(rf"<{tag}\b[^>]*(?:/>|>[\s\S]*?</{tag}>)", sheet_xml)
            fragment = match.group() if match else ""
            settings[tag] = re.sub(r'\s+r:id="[^"]*"', "", fragment)
    payload = {"base64": base64.b64encode(content).decode("ascii"), "printSettings": settings}
    (APP / "TemplateData.gs").write_text(
        "/** Bundled Tamplate_PL.xlsx. Regenerate with tools/build_package.py. */\n"
        "var PL_TEMPLATE = " + json.dumps(payload, ensure_ascii=True, separators=(",", ":")) + ";\n",
        encoding="utf-8",
    )


def build_zip():
    manifest = json.loads((APP / "appsscript.json").read_text(encoding="utf-8"))
    services = manifest.get("dependencies", {}).get("enabledAdvancedServices", [])
    for field in ("userSymbol", "serviceId"):
        identifiers = [service[field] for service in services]
        if len(identifiers) != len(set(identifiers)):
            raise ValueError(f"Duplicate advanced service {field} in appsscript.json")
    output = ROOT / "dist" / "PL_Google_Apps_Script.zip"
    output.parent.mkdir(exist_ok=True)
    names = ["Code.gs", "Core.gs", "TemplateData.gs", "Index.html", "appsscript.json", "README.md"]
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.write(APP / name, name)
    print(f"Deployment package: {output} ({output.stat().st_size:,} bytes)")


if __name__ == "__main__":
    build_template()
    build_zip()

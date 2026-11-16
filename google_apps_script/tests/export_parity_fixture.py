"""Emit an ephemeral parity fixture from the existing Python app and source XLSX.

The GL's financial records are read locally at test time; they are not copied
into the deployable Apps Script package. This module never changes a workbook.
"""

from __future__ import annotations

from dataclasses import asdict
from datetime import date, datetime
from io import BytesIO
import json
import math
from pathlib import Path
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import pandas as pd
from openpyxl import Workbook, load_workbook

from parser import MONTH_ALIASES, parse_period
from pl_builder import DISPLAY_ROWS, apply_edits, build_workbook, default_justifications, preview_pl, read_gl


def native(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and math.isnan(value):
        return None
    if hasattr(value, "item"):
        return native(value.item())
    if isinstance(value, dict):
        return {key: native(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [native(item) for item in value]
    return value


def records(gl):
    return [native({
        "row": entry["Baris Excel"], "project": entry["Project"],
        "account": entry["GL Account"], "postingDate": entry["Posting Date"],
        "amount": entry["In profit center local currency"], "text": entry["Text"],
        "textMonth": entry["Bulan (Text)"], "textYear": entry["Tahun (Text)"],
        "period": entry["Periode PL"], "notes": entry["Catatan"],
    }) for _, entry in gl.records.iterrows()]


def preview(gl, template, year, month):
    result = preview_pl(gl, template, year, month)
    columns = list(result.columns)[2:]
    return native({
        "columns": columns,
        "rows": [{"row": entry["Baris"], "label": entry["ITEMS"],
                  "values": [entry[key] for key in columns]}
                 for _, entry in result.iterrows()],
        "evaluationRate": result.set_index("Baris").at[37, f"TOTAL YTD {year}"],
        "justifications": default_justifications(result, year),
    })


def workbook_rows(contents):
    worksheet = load_workbook(BytesIO(contents)).worksheets[0]
    return native(list(worksheet.iter_rows(values_only=True)))


def fixture(contents, template, year, month, name, edits=None):
    gl = read_gl(contents, year, month)
    result = {
        "name": name, "rows": workbook_rows(contents), "reportYear": year,
        "ytdMonth": month, "project": gl.project, "headers": gl.headers,
        "records": records(gl), "preview": preview(gl, template, year, month),
    }
    if edits:
        transitions = []
        for step in edits:
            apply_edits(gl, pd.DataFrame([{
                "Baris Excel": edit["row"], "Bulan (Text)": edit["textMonth"],
                "Tahun (Text)": edit["textYear"],
            } for edit in step]), year)
            transitions.append({
                "edits": step, "records": records(gl),
                "preview": preview(gl, template, year, month),
            })
        result["transitions"] = transitions
    return result


def synthetic_contents():
    wb = Workbook()
    ws = wb.active
    ws.title = "SAP 'Test'"
    # Deliberately shuffled columns and a blank source row test stable row IDs.
    ws.append(["Text", "Posting Date", "Project", "In profit center local currency", "GL Account"])
    ws.append(["Jan 2026", datetime(2026, 1, 15), "Test", -100, "OPERATING REVENUE"])
    ws.append(["Jan 2026", datetime(2026, 3, 1), "Test", -50, "OPERATING REVENUE"])
    ws.append(["JAN-MAR 26", datetime(2026, 2, 10), "Test", 10, "RENT"])
    ws.append([None] * 5)
    ws.append(["JUL-DES 25", datetime(2026, 3, 10), "Test", 20, "RENT"])
    ws.append(["MAR", datetime(2026, 2, 12), "Test", 30, "RENT"])
    ws.append(["FEB", datetime(2026, 1, 12), "Test", 15, "RENT"])
    ws.append(["SEWA MT PTS JAMBI DES 24", datetime(2026, 1, 20), "Test", 40, "RENT"])
    ws.append(["SEWA MT PTS JAMBI NOV 24", datetime(2026, 1, 21), "Test", 50, "RENT"])
    ws.append(["GAJI DES 24", datetime(2026, 1, 22), "Test", -5, "DIRECT LABOR"])
    ws.append(["UPAH DES 24", datetime(2026, 2, 22), "Test", 6, "SALARY AND COMPENSATION"])
    ws.append(["DEP DES 24", datetime(2026, 2, 22), "Test", 7, "DEPRECIATION & AMORTIZATION"])
    ws.append(["Fuel Jan 26", datetime(2026, 1, 22), "Test", -3100000, "FUEL"])
    # Gross profit just below overhead exercises the monthly max(0, GP-overhead).
    ws.append(["Jan 26", datetime(2026, 1, 22), "Test", 3100099.5, "RENT"])
    ws.append(["Periode 2025", datetime(2025, 12, 22), "Test", 9999, "RENT"])
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def parser_cases():
    cases = [
        ("Parkir bandara 22 apr", "2026-05-17", "RENT", 2026, 12),
        ("SEWA MT 16 KL DES 25", "2026-04-21", "RENT", 2026, 12),
        ("ACCRUE CC OPS AVIASI P.LYG 01-30 APR'26", "2026-05-17", "RENT", 2026, 12),
        ("BIAYA JASA PPH 21 OS LAM PER JAN - MAR 26- FEE", "2026-06-11", "RENT", 2026, 12),
        ("SEWA CCTV AVIASI P LAYANG JUL-DES 25", "2026-04-20", "RENT", 2026, 12),
        ("BIAYA KAK OS LAM AMT PERIODE 2025 - FEE (Technical Fee)", "2026-01-14", "RENT", 2026, 12),
        ("Parkir bandara 26 april", "2026-05-17", "RENT", 2026, 12),
        ("Aqua isi ulang", "2026-03-25", "RENT", 2026, 12),
        ("Extrafooding satgas 21", "2026-03-25", "RENT", 2026, 12),
        ("ADD ACCR SEWA MT AVIASI P. LAYANG 25", "2026-02-27", "RENT", 2026, 12),
        ("Halim18 Feb 25", "2026-05-01", "RENT", 2026, 12),
        ("ACCR TAMBAHAN AGU 20", "2026-08-30", "RENT", 2026, 12),
        ("SEWA DES 24 INVOICE 2025", "2026-01-15", "RENT", 2026, 12),
        ("Biaya Jan 2026", "2026-02-01", "RENT", 2027, 12),
        ("Biaya JAN - MAR '26", "2026-06-11", "RENT", 2026, 12),
        (None, "2026-02-01", "RENT", 2026, 2),
        ('  SEWA\t“DES” "24"  ', "2026-01-15", "RENT", 2026, 2),
        ("BIAYA PERIODE 2026", "2026-07-14", "RENT", 2026, 12),
        ("BIAYA PERIODE 2025", "2026-07-14", "RENT", 2026, 12),
    ]
    for aliases in MONTH_ALIASES.values():
        for alias in aliases:
            for cutoff, suffix in ((2, ""), (12, ""), (2, " 24"), (2, "-24"), (2, "'23"), (2, "/25"), (2, " 2026")):
                cases.append((f"BIAYA {alias.upper()}{suffix}", "2026-06-15", "RENT", 2026, cutoff))
    for account in ("Direct Labor", "Salary and Compensation", "Depreciation & Amortization"):
        for text in ("BIAYA SEP 25", "BIAYA PERIODE 2026", "GAJI DES 24", "MAR"):
            cases.append((text, "2026-01-14", account, 2026, 2))
    result = []
    for text, posting, account, year, cutoff in cases:
        parsed = parse_period(text, date.fromisoformat(posting), account, year, cutoff)
        expected = asdict(parsed)
        result.append({
            "args": [text, posting, account, year, cutoff],
            "expected": {
                "textMonth": expected["text_month"], "textYear": expected["text_year"],
                "periodMonth": expected["period_month"], "periodYear": expected["period_year"],
                "carryOverYear": expected["carry_over_year"], "period": parsed.period_label,
                "notes": "; ".join(parsed.notes),
            },
        })
    return result


def template_xml(contents):
    with ZipFile(BytesIO(contents)) as archive:
        workbook = archive.read("xl/workbook.xml").decode("utf-8")
        root = ET.fromstring(workbook)
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        pl = next(sheet for sheet in root.find("s:sheets", ns) if sheet.attrib["name"] == "PL")
        relationship = pl.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(item.attrib["Target"] for item in relations if item.attrib["Id"] == relationship)
        target = target.lstrip("/") if target.startswith("/") else "xl/" + target
        return {
            "workbook": workbook, "sheet": archive.read(target).decode("utf-8"),
            "styles": archive.read("xl/styles.xml").decode("utf-8"),
            "plIndex": list(root.find("s:sheets", ns)).index(pl),
        }


def main():
    template = (ROOT / "Tamplate_PL.xlsx").read_bytes()
    template_ws = load_workbook(BytesIO(template))["PL"]
    labels = {row: str(template_ws.cell(row, 2).value or "").strip() for row in DISPLAY_ROWS}
    source = (ROOT / "examples" / "GL Aviasi Pulau Layang.xlsx").read_bytes()
    synthetic = synthetic_contents()
    edits = [
        [{"row": 14, "textMonth": 2, "textYear": 2026}],
        [{"row": 14, "textMonth": None, "textYear": 2026}],
        [{"row": 14, "textMonth": 11, "textYear": 2024}],
        [{"row": 11, "textMonth": 12, "textYear": 2024}],
        [{"row": 7, "textMonth": None, "textYear": None}],
    ]
    payload = {
        "labels": labels, "parseCases": parser_cases(),
        "templateXml": template_xml(template),
        "workbooks": [
            fixture(source, template, 2026, 12, "actual GL full year"),
            fixture(source, template, 2026, 2, "actual GL February YTD"),
            fixture(synthetic, template, 2026, 2, "synthetic February YTD", edits),
            fixture(synthetic, template, 2026, 12, "synthetic full year"),
        ],
    }
    # Formula expectations come from the original Python XLSX writer, with the
    # few documented preview/workbook differences checked explicitly in Node.
    gl = read_gl(synthetic, 2026, 2)
    result = load_workbook(BytesIO(build_workbook(gl, template, 2026, ytd_month=2)))["PL"]
    payload["pythonFormulaMatrix"] = {
        "headers": gl.headers, "sheetName": gl.sheet_name,
        "values": [[result.cell(row, col).value for col in range(3, 19)] for row in range(4, 49)],
    }
    print(json.dumps(native(payload), ensure_ascii=True, allow_nan=False, separators=(",", ":")))


if __name__ == "__main__":
    main()

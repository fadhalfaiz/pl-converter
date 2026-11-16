"""Read a GL workbook, calculate the preview, and write the formula-based PL."""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
import shutil
import re
import subprocess
import tempfile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, quote_sheetname, range_boundaries
from openpyxl.workbook.properties import CalcProperties

from parser import effective_period, parse_period


REQUIRED_HEADERS = (
    "Project", "GL Account", "Posting Date",
    "In profit center local currency", "Text",
)
ITEM_ROWS = (4, 7, 8, 9, *range(11, 33))
MONTH_NAMES = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DES")
DISPLAY_ROWS = (4, 6, 7, 8, 9, 10, *range(11, 35), 36, 37, 39, 41, 42, 43, 44, 46, 48)


@dataclass
class GLData:
    source_bytes: bytes
    sheet_name: str
    headers: dict[str, int]
    project: str
    records: pd.DataFrame


def _read_workbook(contents: bytes):
    try:
        return load_workbook(BytesIO(contents))
    except Exception as exc:
        raise ValueError("File tidak dapat dibaca sebagai workbook .xlsx.") from exc


def read_gl(contents: bytes, report_year: int, ytd_month: int = 12) -> GLData:
    wb = _read_workbook(contents)
    ws = wb.worksheets[0]
    headers = {
        str(ws.cell(1, col).value).strip().casefold(): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(1, col).value is not None
    }
    missing = [name for name in REQUIRED_HEADERS if name.casefold() not in headers]
    if missing:
        raise ValueError("Kolom GL tidak ditemukan: " + ", ".join(missing))
    if ws.max_column != 5:
        raise ValueError("File GL harus memiliki tepat lima kolom sumber agar kolom bantu F:I tersedia.")
    mapped = {name: headers[name.casefold()] for name in REQUIRED_HEADERS}
    records = []
    projects = set()
    for row in range(2, ws.max_row + 1):
        vals = {name: ws.cell(row, col).value for name, col in mapped.items()}
        if all(value is None for value in vals.values()):
            continue
        project = str(vals["Project"] or "").strip()
        if not project:
            raise ValueError(f"Project kosong pada baris {row}.")
        projects.add(project)
        if not isinstance(vals["Posting Date"], (date, datetime)):
            raise ValueError(f"Posting Date bukan tanggal Excel pada baris {row}.")
        if not isinstance(vals["In profit center local currency"], (int, float)):
            raise ValueError(f"Nilai GL bukan angka pada baris {row}.")
        if not str(vals["GL Account"] or "").strip():
            raise ValueError(f"GL Account kosong pada baris {row}.")
        parsed = parse_period(
            vals["Text"], vals["Posting Date"], vals["GL Account"], report_year, ytd_month
        )
        records.append({
            "Baris Excel": row,
            **vals,
            "Bulan (Text)": parsed.text_month,
            "Tahun (Text)": parsed.text_year,
            "Periode PL": parsed.period_label,
            "Catatan": "; ".join(parsed.notes),
        })
    if not records:
        raise ValueError("File GL tidak memiliki baris data.")
    if len(projects) != 1:
        raise ValueError("File GL memuat lebih dari satu Project. Unggah satu proyek per laporan.")
    return GLData(contents, ws.title, mapped, projects.pop(), pd.DataFrame.from_records(records))


def apply_edits(gl: GLData, edits: pd.DataFrame, report_year: int) -> None:
    """Apply F/G changes from the review tab, keyed by the stable Excel row."""
    current = gl.records.set_index("Baris Excel")
    for _, edit in edits.iterrows():
        row = int(edit["Baris Excel"])
        if row not in current.index:
            continue
        edited_values = {}
        for key, low, high in (("Bulan (Text)", 1, 12), ("Tahun (Text)", 1900, 9999)):
            value = edit[key]
            if pd.isna(value) or value == "":
                value = None
            else:
                try:
                    if float(value) != int(value):
                        raise ValueError()
                    value = int(value)
                except (TypeError, ValueError) as exc:
                    raise ValueError(f"{key} tidak valid pada baris {row}.") from exc
                if not low <= value <= high:
                    raise ValueError(f"{key} di luar rentang pada baris {row}.")
            current.at[row, key] = value
            edited_values[key] = value
        month, year, carry = effective_period(
            current.at[row, "Posting Date"], current.at[row, "GL Account"], report_year,
            edited_values["Bulan (Text)"], edited_values["Tahun (Text)"],
        )
        current.at[row, "Periode PL"] = f"CO {carry}" if carry else f"{month:02d}-{year}"
    gl.records = current.reset_index()


def _labels(template_bytes: bytes) -> dict[int, str]:
    wb = _read_workbook(template_bytes)
    if "PL" not in wb:
        raise ValueError('Template harus memiliki sheet bernama "PL".')
    return {row: str(wb["PL"].cell(row, 2).value or "").strip() for row in DISPLAY_ROWS}


def preview_pl(
    gl: GLData, template_bytes: bytes, report_year: int, ytd_month: int = 12
) -> pd.DataFrame:
    if not 1 <= ytd_month <= 12:
        raise ValueError("Bulan YTD harus antara Januari dan Desember.")
    labels = _labels(template_bytes)
    accounts = set(gl.records["GL Account"].astype(str).str.strip())
    missing = sorted(accounts - {labels[row] for row in ITEM_ROWS})
    if missing:
        raise ValueError("GL Account tidak ada di template PL: " + ", ".join(missing))

    data = gl.records.copy()
    posting_dates = pd.to_datetime(data["Posting Date"])
    data = data.loc[
        posting_dates.dt.year.eq(report_year) & posting_dates.dt.month.le(ytd_month)
    ].copy()
    data["Account"] = data["GL Account"].astype(str).str.strip()
    data["Amount"] = pd.to_numeric(data["In profit center local currency"])
    values: dict[int, dict[str, float]] = {}
    columns = [f"Carry Over {report_year-1}", *MONTH_NAMES, f"TOTAL YTD {report_year}", "TOTAL SAP", "SELISIH"]
    for row in ITEM_ROWS:
        account = labels[row]
        subset = data[data["Account"] == account]
        sign = -1 if row == 4 else 1
        vals = {name: 0 for name in columns}
        vals[columns[0]] = sign * subset.loc[subset["Periode PL"] == f"CO {report_year-1}", "Amount"].sum()
        for month, name in enumerate(MONTH_NAMES, 1):
            if month <= ytd_month:
                vals[name] = sign * subset.loc[subset["Periode PL"] == f"{month:02d}-{report_year}", "Amount"].sum()
        vals[columns[13]] = sum(vals[name] for name in MONTH_NAMES)
        vals["TOTAL SAP"] = sign * subset["Amount"].sum()
        vals["SELISIH"] = vals[columns[13]] - vals["TOTAL SAP"]
        values[row] = vals

    for row in (6, 10, 42, 43):
        values[row] = {name: 0 for name in columns}
    for col in columns:
        values.setdefault(33, {})[col] = sum(values[r][col] for r in range(11, 33))
        values.setdefault(34, {})[col] = sum(values[r][col] for r in (7, 8, 9, *range(11, 33)))
        values.setdefault(36, {})[col] = values[4][col] - values[34][col]
        values.setdefault(37, {})[col] = values[36][col] / values[4][col] if values[4][col] else 0
        if col == "SELISIH":
            values[37][col] = None
        values.setdefault(39, {})[col] = 0 if col == columns[0] else 0.01 * values[4][col]
        if col == "SELISIH":
            values[39][col] = values[39]["TOTAL SAP"] - values[39][columns[13]]
        values.setdefault(41, {})[col] = (
            0 if col == columns[0] else
            max(0, values[36][col] - values[39][col]) if col in MONTH_NAMES else
            values[36][col] - values[39][col] if col != "SELISIH" else
            values[41][columns[13]] - values[41]["TOTAL SAP"]
        )
        values.setdefault(44, {})[col] = values[41][col] + values[42][col] + values[43][col]
        values.setdefault(46, {})[col] = 0 if col == columns[0] else max(0, 0.22 * values[41][col])
        if col == "SELISIH":
            values[46][col] = values[46]["TOTAL SAP"] - values[46][columns[13]]
        values.setdefault(48, {})[col] = values[44][col] - values[46][col]
        if col == "SELISIH":
            values[48][col] = values[48][columns[13]] - values[48]["TOTAL SAP"]
    return pd.DataFrame([{"Baris": row, "ITEMS": labels[row], **values[row]} for row in DISPLAY_ROWS])


def default_justifications(preview: pd.DataFrame, report_year: int) -> list[str]:
    result = []
    for _, entry in preview[preview["Baris"].isin(ITEM_ROWS)].iterrows():
        amount = entry["SELISIH"]
        if entry[f"Carry Over {report_year-1}"] == 0:
            continue
        name = "Revenue" if entry["Baris"] == 4 else entry["ITEMS"].title()
        size = f"{abs(amount):,.0f}".replace(",", ".")
        direction = "lebih kecil" if amount < 0 else "lebih besar"
        result.append(
            f"{len(result)+1}. {name}, nilai realisasi {direction} sebesar {size} "
            f"dikarenakan adanya Carry Over {report_year-1}"
        )
    return result


def _copy_source_sheet(source, destination) -> None:
    for row in source:
        for old in row:
            new = destination[old.coordinate]
            new.value = old.value
            if old.has_style:
                new._style = copy(old._style)
            if old.number_format:
                new.number_format = old.number_format
            if old.hyperlink:
                new._hyperlink = copy(old.hyperlink)
            if old.comment:
                new.comment = copy(old.comment)
    for key, dim in source.column_dimensions.items():
        destination.column_dimensions[key] = copy(dim)
    for key, dim in source.row_dimensions.items():
        destination.row_dimensions[key] = copy(dim)
    for merged in source.merged_cells.ranges:
        destination.merge_cells(str(merged))
    destination.sheet_format = copy(source.sheet_format)
    destination.views = copy(source.views)
    destination.print_options = copy(source.print_options)
    destination.sheet_properties = copy(source.sheet_properties)


def _numeric_display_format(number_format: str) -> str:
    """Space accounting minus signs from values and display zeros as a dash."""
    sections = number_format.split(";")
    positive = sections[0]
    zero = '_-* "-"??_-'
    if positive == "General":
        return f"#,##0;\\-* #,##0_-;{zero}"
    if "%" in positive:
        negative = r"\-* 0.0%_-" if "0.0%" in positive else r"\-* 0%_-"
    elif "#,##0" in positive:
        negative = r"\-* #,##0_-"
    else:
        return number_format
    if len(sections) == 1:
        sections.append(negative)
    else:
        sections[1] = negative
    if len(sections) == 2:
        sections.append(zero)
    else:
        sections[2] = zero
    return ";".join(sections)


def _justification_formula(line: str, account_row: int) -> str:
    ref = f"$R${account_row}"
    # Format the absolute difference with Indonesian thousands separators.
    amount = f'SUBSTITUTE(SUBSTITUTE(TEXT(ABS({ref}),"#,##0"),",",".")," ",".")'
    quote = lambda text: '"' + text.replace('"', '""') + '"'
    match = re.search(r"sebesar\s+(?:Rp\.?\s*)?-?\s*\d[\d.,]*", line, re.IGNORECASE)
    if match:
        prefix, suffix = line[:match.start()], line[match.end():]
        parts = re.split(r"(lebih kecil|lebih besar)", prefix)
        tokens = [
            f'IF({ref}<0,"lebih kecil","lebih besar")'
            if part in ("lebih kecil", "lebih besar") else quote(part)
            for part in parts
        ]
        return "=" + "&".join(tokens + [quote("sebesar Rp "), amount, quote(suffix)])
    return "=" + quote(line + " — Selisih Rp ") + f'&IF({ref}<0,"− ","")&' + amount


def _set_justifications(ws, lines: list[str], account_rows: list[int]) -> None:
    for row in range(54, 61):
        ws.cell(row, 2).value = None
    extra = max(0, len(lines) + 2 - 7)
    if extra:
        signature_merges = [str(rng) for rng in ws.merged_cells.ranges if rng.min_row >= 61]
        for rng in signature_merges:
            ws.unmerge_cells(rng)
        ws.move_range("B61:Z224", rows=extra)
        for rng in signature_merges:
            min_col, min_row, max_col, max_row = range_boundaries(rng)
            ws.merge_cells(start_row=min_row + extra, start_column=min_col,
                           end_row=max_row + extra, end_column=max_col)
        for row in range(224, 60, -1):
            if ws.row_dimensions[row].height is not None:
                ws.row_dimensions[row + extra].height = ws.row_dimensions[row].height
        if ws.print_area:
            area = str(ws.print_area).split("!", 1)[-1]
            min_col, min_row, max_col, max_row = range_boundaries(area)
            ws.print_area = (
                f"{get_column_letter(min_col)}{min_row}:"
                f"{get_column_letter(max_col)}{max_row + extra}"
            )
    style = copy(ws["B54"]._style)
    for index, line in enumerate(lines, 54):
        cell = ws.cell(index, 2)
        point = index - 54
        cell.value = _justification_formula(line, account_rows[point]) if point < len(account_rows) else line
        cell._style = copy(style)
    evaluation_row = 54 + len(lines)
    ws.merge_cells(start_row=evaluation_row, start_column=2, end_row=evaluation_row, end_column=18)
    ws.merge_cells(start_row=evaluation_row + 1, start_column=2, end_row=evaluation_row + 1, end_column=18)
    title = ws.cell(evaluation_row, 2)
    title.value = "Hasil Evaluasi :"
    title._style = copy(ws["B52"]._style)
    sentence = ws.cell(evaluation_row + 1, 2)
    # Build decimal digits explicitly: TEXT's decimal mask depends on Excel's locale.
    rate = "IFERROR(P48/P4,0)"
    hundredths = f"ROUND(ABS({rate})*10000,0)"
    sentence.value = (
        '="Nilai NPAT setelah dilakukan cleansing sebesar "'
        f'&IF({rate}<0,"− ","")'
        f'&INT({hundredths}/100)&","&RIGHT("00"&MOD({hundredths},100),2)&"%"'
    )
    sentence._style = copy(style)


def build_workbook(
    gl: GLData, template_bytes: bytes, report_year: int,
    justifications: list[str] | None = None,
    ytd_month: int = 12,
    creator_name: str | None = None,
    creator_role: str | None = None,
    reviewer_name: str | None = None,
    reviewer_role: str | None = "Manager TOPM Dept.",
    approver_name: str | None = None,
    approver_role: str | None = None,
) -> bytes:
    if not 1 <= ytd_month <= 12:
        raise ValueError("Bulan YTD harus antara Januari dan Desember.")
    source_wb = _read_workbook(gl.source_bytes)
    source = source_wb.worksheets[0]
    wb = _read_workbook(template_bytes)
    if "PL" not in wb:
        raise ValueError('Template harus memiliki sheet bernama "PL".')
    for name in list(wb.sheetnames):
        if name != "PL":
            del wb[name]
    if gl.sheet_name == "PL":
        raise ValueError('Nama sheet GL "PL" bertabrakan dengan sheet laporan.')
    data = wb.create_sheet(gl.sheet_name, 0)
    _copy_source_sheet(source, data)
    amount_col = gl.headers["In profit center local currency"]
    for row in range(2, source.max_row + 1):
        cell = data.cell(row, amount_col)
        if isinstance(cell.value, (int, float)):
            cell.number_format = _numeric_display_format(cell.number_format)
    header_fill = PatternFill("solid", fgColor="FFF2CC")
    header_font = Font(name="Calibri Light", size=12, bold=True)
    for col, name in enumerate(("Bulan (Text)", "Tahun (Text)", "Periode PL", "Catatan"), 6):
        cell = data.cell(1, col, name)
        cell.fill = header_fill
        cell.font = header_font
    for col in range(1, 6):
        data.cell(1, col).fill = header_fill
        data.cell(1, col).font = header_font
    data.auto_filter.ref = f"A1:I{source.max_row}"
    data.freeze_panes = "A2"
    for col, width in {"F": 17, "G": 17, "H": 17, "I": 80}.items():
        data.column_dimensions[col].width = width
    blue_font = Font(name="Calibri Light", size=12, color="0000FF")
    account_col = get_column_letter(gl.headers["GL Account"])
    posting_col = get_column_letter(gl.headers["Posting Date"])
    for _, record in gl.records.iterrows():
        row = int(record["Baris Excel"])
        for col, key in ((6, "Bulan (Text)"), (7, "Tahun (Text)")):
            value = record[key]
            cell = data.cell(row, col)
            cell.value = None if pd.isna(value) else int(value)
            cell.font = blue_font
        a = f"${account_col}{row}"
        p = f"${posting_col}{row}"
        data.cell(row, 8).value = (
            f'=IF(OR({a}="DIRECT LABOR",{a}="SALARY AND COMPENSATION",'
            f'{a}="DEPRECIATION & AMORTIZATION"),DATE(YEAR({p}),MONTH({p}),1),'
            f'IF(AND($G{row}<>"",$G{row}<PL!$C$2),"CO "&$G{row},'
            f'IF(AND($G{row}<>"",$F{row}<>""),DATE($G{row},$F{row},1),'
            f'DATE(YEAR({p}),MONTH({p}),1))))'
        )
        data.cell(row, 8).number_format = "mmm-yy"
        data.cell(row, 9).value = record["Catatan"] or None

    pl = wb["PL"]
    pl["B1"] = f"PL REALISASI {gl.project.upper()}"
    pl["B2"] = "TAHUN"
    pl["C2"] = report_year
    pl["B2"].font = Font(name="Calibri Light", size=12)
    pl["C2"].font = Font(name="Calibri Light", size=12, color="0000FF")
    pl["C2"].fill = header_fill
    pl["AA1"] = "BULAN YTD"
    pl["AA2"] = ytd_month
    pl.column_dimensions["AA"].hidden = True
    pl["C3"] = '="Carry over "&($C$2-1)'
    pl["P3"] = '="TOTAL YTD "&$C$2'
    pl["Q3"] = "TOTAL SAP"
    pl["R3"] = "SELISIH"
    for month, name in enumerate(MONTH_NAMES, 4):
        pl.cell(3, month).value = name
    sheet = quote_sheetname(data.title)
    amt = f"{sheet}!${get_column_letter(gl.headers['In profit center local currency'])}:${get_column_letter(gl.headers['In profit center local currency'])}"
    acc = f"{sheet}!${account_col}:${account_col}"
    posting = f"{sheet}!${posting_col}:${posting_col}"
    period = f"{sheet}!$H:$H"
    posting_window = (
        f'{posting},">="&DATE($C$2,1,1),'
        f'{posting},"<"&DATE($C$2,$AA$2+1,1)'
    )
    for row in ITEM_ROWS:
        sign = "-" if row == 4 else ""
        pl.cell(row, 3).value = (
            f'={sign}SUMIFS({amt},{acc},$B{row},{period},"CO "&($C$2-1),'
            f'{posting_window})'
        )
        for month in range(1, 13):
            pl.cell(row, month + 3).value = (
                f'=IF({month}<=$AA$2,{sign}SUMIFS({amt},{acc},$B{row},'
                f'{period},">="&DATE($C$2,{month},1),'
                f'{period},"<"&DATE($C$2,{month+1},1),{posting_window}),0)'
            )
        pl.cell(row, 16).value = f"=SUM(D{row}:O{row})"
        pl.cell(row, 17).value = f'={sign}SUMIFS({amt},{acc},$B{row},{posting_window})'
        pl.cell(row, 18).value = f"=P{row}-Q{row}"
    for col_num in (3, *range(4, 19)):
        col = get_column_letter(col_num)
        pl[f"{col}33"] = f"=SUM({col}11:{col}32)"
        pl[f"{col}34"] = f"=SUM({col}7:{col}9,{col}11:{col}32)"
        pl[f"{col}36"] = f"={col}4-{col}34" if col_num != 18 else "=P36-Q36"
        if col_num != 18:
            pl[f"{col}37"] = f"=IFERROR({col}36/{col}4,0)"
        if col_num >= 4:
            pl[f"{col}39"] = f"=1%*{col}4" if col_num != 18 else "=Q39-P39"
            pl[f"{col}41"] = f"=IF({col}36<0,0,{col}36-{col}39)" if col_num <= 15 else f"={col}36-{col}39" if col_num <= 17 else "=P41-Q41"
            if col_num <= 17:
                pl[f"{col}42"] = "=0"
                pl[f"{col}43"] = "=0"
            else:
                pl[f"{col}42"] = "=P42-Q42"
                pl[f"{col}43"] = "=P43-Q43"
            pl[f"{col}44"] = f"=SUM({col}41:{col}43)" if col_num <= 17 else "=P44-Q44"
            pl[f"{col}46"] = f"=IF({col}41<0,0,22%*{col}41)" if col_num <= 17 else "=Q46-P46"
            pl[f"{col}48"] = f"={col}44-{col}46" if col_num <= 17 else "=P48-Q48"
    for row in (33, 34, 36):
        pl[f"R{row}"] = f"=P{row}-Q{row}"
    for row in (6, 10):
        for col in range(3, 19):
            pl.cell(row, col).value = None
    for row in range(4, 49):
        for col in range(3, 19):
            cell = pl.cell(row, col)
            if cell.data_type == "f":
                cell.number_format = _numeric_display_format(cell.number_format)
    preview = preview_pl(gl, template_bytes, report_year, ytd_month)
    account_rows = preview.loc[preview["Baris"].isin(ITEM_ROWS) & preview[f"Carry Over {report_year-1}"].ne(0), "Baris"].tolist()
    _set_justifications(pl, justifications if justifications is not None else default_justifications(preview, report_year), account_rows)
    if creator_name is not None or creator_role is not None:
        creator_row = next(
            (row for row in range(1, pl.max_row + 1) if pl.cell(row, 9).value == "Dibuat oleh,"),
            None,
        )
        if creator_row is None:
            raise ValueError('Bagian "Dibuat oleh" tidak ditemukan di template PL.')
        for offset, value in ((5, creator_name), (6, creator_role)):
            if value is not None:
                cell = pl.cell(creator_row + offset, 9)
                cell.value = value.strip()
                cell.data_type = "s"
    if reviewer_name is not None or reviewer_role is not None:
        reviewer_row = next(
            (row for row in range(1, pl.max_row + 1) if pl.cell(row, 6).value == "Mengetahui,"),
            None,
        )
        if reviewer_row is None:
            raise ValueError('Bagian "Mengetahui" tidak ditemukan di template PL.')
        for offset, value in ((5, reviewer_name), (6, reviewer_role)):
            if value is not None:
                cell = pl.cell(reviewer_row + offset, 6)
                cell.value = value.strip()
                cell.data_type = "s"
    if approver_name is not None or approver_role is not None:
        approver_row = next(
            (row for row in range(1, pl.max_row + 1) if pl.cell(row, 3).value == "Menyetujui,"),
            None,
        )
        if approver_row is None:
            raise ValueError('Bagian "Menyetujui" tidak ditemukan di template PL.')
        for offset, value in ((5, approver_name), (6, approver_role)):
            if value is not None:
                cell = pl.cell(approver_row + offset, 3)
                cell.value = value.strip()
                cell.data_type = "s"
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def recalculate_with_libreoffice(contents: bytes) -> tuple[bytes, str]:
    """Recalculate when LibreOffice is installed, and scan cached formula errors."""
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if not executable:
        for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if Path(candidate).exists():
                executable = candidate
                break
    if not executable:
        return contents, "LibreOffice tidak tersedia; Excel akan menghitung rumus saat file dibuka."
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "input" / "report.xlsx"
        target_dir = root / "output"
        source.parent.mkdir()
        target_dir.mkdir()
        source.write_bytes(contents)
        command = [executable, "-env:UserInstallation=" + (root / "profile").as_uri(),
                   "--headless", "--convert-to", "xlsx", "--outdir", str(target_dir), str(source)]
        run = subprocess.run(command, capture_output=True, text=True, timeout=180, check=False)
        target = target_dir / "report.xlsx"
        if run.returncode != 0 or not target.exists():
            raise RuntimeError("LibreOffice gagal menghitung workbook: " + (run.stderr or run.stdout).strip())
        result = target.read_bytes()
    cached = load_workbook(BytesIO(result), data_only=True)
    errors = []
    for ws in cached:
        for row in ws:
            for cell in row:
                if cell.data_type == "e" or (isinstance(cell.value, str) and cell.value in {"#NAME?", "#VALUE!", "#REF!"}):
                    errors.append(f"{ws.title}!{cell.coordinate}: {cell.value}")
    if errors:
        raise RuntimeError("Ditemukan error rumus: " + ", ".join(errors[:10]))
    return result, "Rumus dihitung ulang dengan LibreOffice."

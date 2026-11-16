import io
from copy import copy
from pathlib import Path
import unittest

from openpyxl import Workbook, load_workbook

from pl_builder import ITEM_ROWS, apply_edits, build_workbook, preview_pl, read_gl


ROOT = Path(__file__).resolve().parents[1]
EXPECTED = {
    4: (-1414313540, 2283104402, 14789559431, 13375245891),
    9: (261690600, 358560000, 1411166905, 1672857505),
    17: (-1985782797, 514317435, 1953256021, -32526776),
    19: (159552150, 254244630, 2279929700, 2439481850),
    8: (0, 352834711, 2705660339, 2705660339),
    34: (-1616606335, 1610087715, 12126064427, 10516182647),
    36: (202292795, 673016687, 2663495004, 2859063244),
}


class IntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (ROOT / "examples" / "GL Aviasi Pulau Layang.xlsx").read_bytes()
        cls.template = (ROOT / "Tamplate_PL.xlsx").read_bytes()

    def test_acceptance_preview_and_all_accounts(self):
        gl = read_gl(self.source, 2026)
        preview = preview_pl(gl, self.template, 2026).set_index("Baris")
        for row, expected in EXPECTED.items():
            actual = tuple(int(preview.at[row, col]) for col in
                           ("Carry Over 2025", "JAN", "TOTAL YTD 2026", "TOTAL SAP"))
            self.assertEqual(actual, expected, row)
        labels = set(preview.loc[list(ITEM_ROWS), "ITEMS"])
        self.assertEqual(len(set(gl.records["GL Account"])), 19)
        self.assertEqual(set(gl.records["GL Account"]), labels & set(gl.records["GL Account"]))
        self.assertEqual(gl.records.set_index("Baris Excel").at[25, "Periode PL"], "CO 2020")
        for row in ITEM_ROWS:
            self.assertEqual(int(preview.at[row, "SELISIH"]),
                             -int(preview.at[row, "Carry Over 2025"])
                             - (6724555 if row == 27 else 0))

    def test_output_formulas_template_and_signatures(self):
        gl = read_gl(self.source, 2026)
        result = build_workbook(gl, self.template, 2026)
        wb = load_workbook(io.BytesIO(result))
        self.assertEqual(wb.sheetnames, ["Sheet1", "PL"])
        data, pl = wb.worksheets
        source_ws = load_workbook(io.BytesIO(self.source)).worksheets[0]
        for row in (1, 2, 600, 1235):
            self.assertEqual([data.cell(row, col).value for col in range(1, 6)],
                             [source_ws.cell(row, col).value for col in range(1, 6)])
        self.assertEqual(data["D2"].number_format, '#,##0;\\-* #,##0_-;_-* "-"??_-')
        self.assertEqual(data.auto_filter.ref, "A1:I1235")
        self.assertEqual(data.freeze_panes, "A2")
        self.assertTrue(data["H2"].value.startswith("=IF(OR("))
        self.assertEqual(data["G1097"].value, 2025)
        self.assertIn("DIRECT LABOR", data["H1097"].value)
        self.assertEqual(gl.records.set_index("Baris Excel").at[1097, "Periode PL"], "01-2026")
        self.assertEqual(data["I1097"].value, "GL dikecualikan -> ikut posting date")
        self.assertEqual(data["H2"].number_format, "mmm-yy")
        self.assertIn("SUMIFS", pl["D4"].value)
        self.assertIn("-SUMIFS", pl["D4"].value)
        self.assertTrue(pl["Q4"].value.startswith("=-SUMIFS("))
        self.assertEqual(pl["P4"].value, "=SUM(D4:O4)")
        self.assertEqual(pl["R33"].value, "=P33-Q33")
        self.assertEqual(pl["R34"].value, "=P34-Q34")
        self.assertEqual(pl["R36"].value, "=P36-Q36")
        self.assertEqual(pl["O48"].value, "=O44-O46")
        self.assertIn("Carry Over 2025", pl["B62"].value)
        preview = preview_pl(gl, self.template, 2026)
        justification_rows = preview.loc[preview["Baris"].isin(ITEM_ROWS) & preview["Carry Over 2025"].ne(0), "Baris"]
        for output_row, account_row in enumerate(justification_rows, 54):
            self.assertEqual(pl[f"B{output_row}"].data_type, "f")
            self.assertIn(f"$R${account_row}", pl[f"B{output_row}"].value)
        self.assertEqual(pl["B63"].value, "Hasil Evaluasi :")
        self.assertIn('ROUND(ABS(IFERROR(P48/P4,0))*10000,0)', pl["B64"].value)
        self.assertIn('&","&RIGHT("00"&MOD(', pl["B64"].value)
        self.assertIn('IF(IFERROR(P48/P4,0)<0,"− ","")', pl["B64"].value)
        self.assertTrue(pl["B64"].value.startswith('="Nilai NPAT setelah dilakukan cleansing sebesar "'))
        signature_row = next(row for row in range(61, 75) if pl[f"C{row}"].value == "Menyetujui,")
        self.assertEqual(signature_row, 65)
        self.assertIn(f"Q{signature_row}:R{signature_row}", {str(r) for r in pl.merged_cells.ranges})
        self.assertEqual(pl[f"C{signature_row + 5}"].value, "Rahman Ilahi")
        self.assertEqual(pl[f"C{signature_row + 6}"].value.strip(), "GM Transportation Operation")
        creator_row = next(row for row in range(61, 75) if pl[f"I{row}"].value == "Dibuat oleh,")
        self.assertEqual(pl[f"I{creator_row + 5}"].value, "Fadhlillah Rifqi H")
        self.assertEqual(pl[f"I{creator_row + 6}"].value.strip(), "UH Performance BBM")
        reviewer_row = next(row for row in range(61, 75) if pl[f"F{row}"].value == "Mengetahui,")
        self.assertEqual(pl[f"F{reviewer_row + 5}"].value.strip(), "Rizky Farlyan")
        self.assertEqual(pl[f"F{reviewer_row + 6}"].value, "Manager TOPM Dept.")
        self.assertEqual(pl["C2"].value, 2026)
        self.assertIsNone(pl["M2"].value)
        self.assertIsNone(pl["N2"].value)
        self.assertEqual(pl["AA2"].value, 12)
        self.assertTrue(pl.column_dimensions["AA"].hidden)
        self.assertIn("TOTAL YTD", pl["P3"].value)
        self.assertEqual(pl["C2"].font.color.rgb, "000000FF")
        self.assertEqual(pl["C2"].fill.fgColor.rgb, "00FFF2CC")
        original_pl = load_workbook(io.BytesIO(self.template))["PL"]
        self.assertEqual(pl["B4"]._style, original_pl["B4"]._style)
        self.assertEqual(copy(pl["D4"].font), copy(original_pl["D4"].font))
        self.assertEqual(copy(pl["D4"].fill), copy(original_pl["D4"].fill))
        self.assertEqual(copy(pl["D4"].border), copy(original_pl["D4"].border))
        self.assertEqual(pl["D4"].number_format.split(";")[0], original_pl["D4"].number_format.split(";")[0])
        self.assertEqual(pl["D4"].number_format.split(";")[1], r"\-* #,##0_-")
        self.assertEqual(pl["C19"].number_format.split(";")[1], pl["D19"].number_format.split(";")[1])
        self.assertNotIn("(", pl["D17"].number_format.split(";")[1])
        self.assertEqual(pl["D17"].number_format.split(";")[2], '_-* "-"??_-')
        self.assertEqual(pl["D37"].number_format, '0%;\\-* 0%_-;_-* "-"??_-')
        self.assertEqual(pl["K9"].number_format, '#,##0;\\-* #,##0_-;_-* "-"??_-')
        self.assertEqual({
            cell.number_format.split(";")[2]
            for row in pl.iter_rows(min_row=4, max_row=48, min_col=3, max_col=18)
            for cell in row if cell.data_type == "f"
        }, {'_-* "-"??_-'})
        self.assertEqual(pl["M2"]._style, original_pl["M2"]._style)
        self.assertEqual(pl["N2"]._style, original_pl["N2"]._style)
        self.assertEqual(str(pl.print_area), "'PL'!$A$1:$S$72")
        self.assertFalse([
            cell.coordinate for row in pl.iter_rows(min_row=4, max_row=48, min_col=3, max_col=18)
            for cell in row if isinstance(cell.value, (int, float))
        ])

    def test_editable_creator_is_written_to_shifted_signature(self):
        gl = read_gl(self.source, 2026)
        result = build_workbook(
            gl, self.template, 2026,
            creator_name="Nama Baru", creator_role="Jabatan Baru",
            reviewer_name="Nama Mengetahui", reviewer_role="Jabatan Mengetahui",
            approver_name="Nama Menyetujui", approver_role="Jabatan Menyetujui",
        )
        pl = load_workbook(io.BytesIO(result))["PL"]
        creator_row = next(row for row in range(61, 75) if pl[f"I{row}"].value == "Dibuat oleh,")
        self.assertEqual(pl[f"I{creator_row + 5}"].value, "Nama Baru")
        self.assertEqual(pl[f"I{creator_row + 6}"].value, "Jabatan Baru")
        reviewer_row = next(row for row in range(61, 75) if pl[f"F{row}"].value == "Mengetahui,")
        self.assertEqual(pl[f"F{reviewer_row + 5}"].value, "Nama Mengetahui")
        self.assertEqual(pl[f"F{reviewer_row + 6}"].value, "Jabatan Mengetahui")
        approver_row = next(row for row in range(61, 75) if pl[f"C{row}"].value == "Menyetujui,")
        self.assertEqual(pl[f"C{approver_row + 5}"].value, "Nama Menyetujui")
        self.assertEqual(pl[f"C{approver_row + 6}"].value, "Jabatan Menyetujui")

    def test_evaluation_follows_short_justification_list(self):
        gl = read_gl(self.source, 2026)
        pl = load_workbook(io.BytesIO(build_workbook(
            gl, self.template, 2026, justifications=["Poin satu", "Poin dua"],
        )))["PL"]
        self.assertIn('Poin satu', pl["B54"].value)
        self.assertIn('$R$4', pl["B54"].value)
        self.assertIn('Poin dua', pl["B55"].value)
        self.assertEqual(pl["B55"].data_type, "f")
        self.assertEqual(pl["B56"].value, "Hasil Evaluasi :")
        self.assertIn("IFERROR(P48/P4,0)", pl["B57"].value)
        self.assertEqual(pl["C61"].value, "Menyetujui,")
        self.assertEqual(str(pl.print_area), "'PL'!$A$1:$S$68")

    def test_review_edit_recalculates_preview(self):
        gl = read_gl(self.source, 2026)
        row = gl.records[gl.records["GL Account"] == "FUEL"].iloc[0]
        edited = gl.records.loc[gl.records["Baris Excel"] == row["Baris Excel"],
                                ["Baris Excel", "Bulan (Text)", "Tahun (Text)"]].copy()
        edited.loc[:, "Bulan (Text)"] = 1
        edited.loc[:, "Tahun (Text)"] = 2026
        original = preview_pl(gl, self.template, 2026).set_index("Baris")
        apply_edits(gl, edited, 2026)
        changed = preview_pl(gl, self.template, 2026).set_index("Baris")
        self.assertEqual(changed.at[19, "JAN"] - original.at[19, "JAN"], row["In profit center local currency"])

    def test_header_mapping(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Different name"
        ws.append(["Text", "Posting Date", "Project", "In profit center local currency", "GL Account"])
        from datetime import datetime
        ws.append(["Jan 2026", datetime(2026, 2, 1), "Test", -100, "OPERATING REVENUE"])
        contents = io.BytesIO()
        wb.save(contents)
        gl = read_gl(contents.getvalue(), 2026)
        self.assertEqual(gl.headers["GL Account"], 5)
        result = load_workbook(io.BytesIO(build_workbook(gl, self.template, 2026)))
        self.assertIn("$B2", result.worksheets[0]["H2"].value)
        self.assertIn("$D:$D", result["PL"]["D4"].value)

    def test_ytd_cutoff_and_range_formula(self):
        from datetime import datetime

        wb = Workbook()
        ws = wb.active
        ws.title = "SAP"
        ws.append(["Project", "GL Account", "Posting Date", "In profit center local currency", "Text"])
        ws.append(["Test", "OPERATING REVENUE", datetime(2026, 1, 15), -100, "Jan 2026"])
        ws.append(["Test", "OPERATING REVENUE", datetime(2026, 3, 1), -50, "Jan 2026"])
        ws.append(["Test", "RENT", datetime(2026, 2, 10), 10, "JAN-MAR 26"])
        ws.append(["Test", "RENT", datetime(2026, 3, 10), 20, "JUL-DES 25"])
        ws.append(["Test", "RENT", datetime(2026, 2, 12), 30, "MAR"])
        ws.append(["Test", "RENT", datetime(2026, 1, 12), 15, "FEB"])
        ws.append(["Test", "RENT", datetime(2026, 1, 20), 40, "SEWA MT PTS JAMBI DES 24"])
        ws.append(["Test", "RENT", datetime(2026, 1, 21), 50, "SEWA MT PTS JAMBI NOV 24"])
        source = io.BytesIO()
        wb.save(source)
        gl = read_gl(source.getvalue(), 2026, ytd_month=2)
        full_year = read_gl(source.getvalue(), 2026)
        self.assertEqual(full_year.records.set_index("Baris Excel").at[6, "Tahun (Text)"], 2026)
        preview = preview_pl(gl, self.template, 2026, ytd_month=2).set_index("Baris")
        self.assertEqual(preview.at[4, "JAN"], 100)
        self.assertEqual(preview.at[4, "TOTAL SAP"], 100)
        self.assertEqual(preview.at[17, "FEB"], 25)
        self.assertEqual(preview.at[17, "MAR"], 0)
        self.assertEqual(preview.at[17, "Carry Over 2025"], 30)
        result = load_workbook(io.BytesIO(build_workbook(gl, self.template, 2026, ytd_month=2)))
        data, pl = result["SAP"], result["PL"]
        self.assertIsNone(data["F4"].value)
        self.assertEqual(data["G4"].value, 2026)
        self.assertEqual((data["F6"].value, data["G6"].value), (3, 2025))
        self.assertEqual((data["F7"].value, data["G7"].value), (2, 2026))
        self.assertEqual((data["F8"].value, data["G8"].value), (12, 2024))
        self.assertEqual((data["F9"].value, data["G9"].value), (11, 2024))
        self.assertEqual(gl.records.set_index("Baris Excel").at[8, "Periode PL"], "CO 2024")
        self.assertEqual(gl.records.set_index("Baris Excel").at[9, "Periode PL"], "CO 2024")
        self.assertIsNone(pl["M2"].value)
        self.assertIsNone(pl["N2"].value)
        self.assertEqual(pl["AA2"].value, 2)
        self.assertTrue(pl.column_dimensions["AA"].hidden)
        self.assertIn("$AA$2", pl["D4"].value)
        self.assertIn("$AA$2+1", pl["Q4"].value)
        self.assertIn("$AA$2+1", pl["C17"].value)


if __name__ == "__main__":
    unittest.main()

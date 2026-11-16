import unittest
from datetime import date

from parser import effective_period, parse_period


class ParserTests(unittest.TestCase):
    def test_required_examples(self):
        cases = [
            ("Parkir bandara 22 apr", "2026-05-17", 4, 2026, None),
            ("SEWA MT 16 KL DES 25", "2026-04-21", None, None, 2025),
            ("ACCRUE CC OPS AVIASI P.LYG 01-30 APR'26", "2026-05-17", 4, 2026, None),
            ("BIAYA JASA PPH 21 OS LAM PER JAN - MAR 26- FEE", "2026-06-11", 6, 2026, None),
            ("SEWA CCTV AVIASI P LAYANG JUL-DES 25", "2026-04-20", None, None, 2025),
            ("BIAYA KAK OS LAM AMT PERIODE 2025 - FEE (Technical Fee)", "2026-01-14", None, None, 2025),
            ("Parkir bandara 26 april", "2026-05-17", 4, 2026, None),
            ("Aqua isi ulang", "2026-03-25", 3, 2026, None),
            ("Extrafooding satgas 21", "2026-03-25", 3, 2026, None),
            ("ADD ACCR SEWA MT AVIASI P. LAYANG 25", "2026-02-27", None, None, 2025),
        ]
        for text, posting, month, year, carry in cases:
            with self.subTest(text=text):
                result = parse_period(text, date.fromisoformat(posting), "RENT", 2026)
                self.assertEqual((result.period_month, result.period_year, result.carry_over_year),
                                 (month, year, carry))

    def test_excluded_accounts_always_follow_posting_date(self):
        for account in ("Direct Labor", "Salary and Compensation", "Depreciation & Amortization"):
            with self.subTest(account=account):
                result = parse_period("BIAYA SEP 25", date(2026, 1, 14), account, 2026)
                self.assertEqual((result.period_month, result.period_year, result.carry_over_year),
                                 (1, 2026, None))
                self.assertEqual((result.text_month, result.text_year), (9, 2025))
                self.assertIn("GL dikecualikan", result.notes[0])
                self.assertEqual(effective_period(date(2026, 1, 14), account, 2026, 9, 2025),
                                 (1, 2026, None))
                current = parse_period("BIAYA PERIODE 2026", date(2026, 1, 14), account, 2026)
                self.assertEqual(current.period_label, "01-2026")
                self.assertEqual(effective_period(date(2026, 1, 14), account, 2026, None, 2026),
                                 (1, 2026, None))

    def test_year_only_and_month_only(self):
        posting = date(2026, 7, 14)
        month_only = parse_period("BIAYA APR", posting, "RENT", 2026)
        self.assertEqual((month_only.text_month, month_only.text_year), (4, 2026))
        self.assertEqual((month_only.period_month, month_only.period_year), (4, 2026))
        year_2026 = parse_period("BIAYA PERIODE 2026", posting, "RENT", 2026)
        self.assertEqual((year_2026.period_month, year_2026.period_year, year_2026.carry_over_year),
                         (7, 2026, None))
        year_2025 = parse_period("BIAYA PERIODE 2025", posting, "RENT", 2026)
        self.assertEqual((year_2025.period_month, year_2025.period_year, year_2025.carry_over_year),
                         (None, None, 2025))

    def test_month_without_year_after_ytd_is_previous_year(self):
        posting = date(2026, 2, 14)
        february = parse_period("BIAYA FEB", posting, "RENT", 2026, ytd_month=2)
        self.assertEqual(february.period_label, "02-2026")
        march = parse_period("BIAYA MAR", posting, "RENT", 2026, ytd_month=2)
        self.assertEqual((march.text_month, march.text_year, march.period_label),
                         (3, 2025, "CO 2025"))
        self.assertIn("tahun dianggap 2025", march.notes[0])
        explicit = parse_period("BIAYA MAR 2026", posting, "RENT", 2026, ytd_month=2)
        self.assertEqual(explicit.period_label, "03-2026")
        excluded = parse_period("BIAYA MAR", posting, "DIRECT LABOR", 2026, ytd_month=2)
        self.assertEqual(excluded.period_label, "02-2026")
        excluded_current = parse_period("BIAYA MAR 2026", posting, "DIRECT LABOR", 2026, ytd_month=2)
        self.assertEqual(excluded_current.period_label, "02-2026")

    def test_month_boundaries_and_short_year(self):
        result = parse_period("Halim18 Feb 25", date(2026, 5, 1), "RENT", 2026)
        self.assertEqual((result.text_month, result.text_year), (2, 2025))
        result = parse_period("Parkir bandara 26 april", date(2026, 5, 1), "RENT", 2026)
        self.assertEqual((result.text_month, result.text_year), (4, 2026))
        self.assertIn("Bulan tanpa tahun", result.notes[0])
        result = parse_period("ACCR TAMBAHAN AGU 20", date(2026, 8, 30), "RENT", 2026)
        self.assertEqual((result.text_month, result.text_year, result.period_label),
                         (8, 2020, "CO 2020"))

    def test_month_followed_by_two_digit_year(self):
        for text, month, year in (
            ("SEWA MT PTS JAMBI DES 24", 12, 2024),
            ("SEWA MT PTS JAMBI NOV 24", 11, 2024),
            ("SEWA DES-24", 12, 2024),
            ("SEWA JAN'23", 1, 2023),
            ("SEWA AGU/25", 8, 2025),
            ("SEWA DES 24 INVOICE 2025", 12, 2024),
        ):
            with self.subTest(text=text):
                result = parse_period(text, date(2026, 1, 15), "RENT", 2026)
                self.assertEqual((result.text_month, result.text_year, result.period_label),
                                 (month, year, f"CO {year}"))
        excluded = parse_period("GAJI DES 24", date(2026, 1, 15), "DIRECT LABOR", 2026)
        self.assertEqual((excluded.text_month, excluded.text_year, excluded.period_label),
                         (12, 2024, "01-2026"))
        prior_year = parse_period("SEWA MT PTS JAMBI DES 24", date(2025, 1, 15), "RENT", 2025)
        self.assertEqual(prior_year.period_label, "CO 2024")

    def test_report_year_controls_carry_over(self):
        result = parse_period("Biaya Jan 2026", date(2026, 2, 1), "RENT", 2027)
        self.assertEqual(result.carry_over_year, 2026)

    def test_month_range_in_report_year_follows_posting_date(self):
        result = parse_period("Biaya JAN - MAR '26", date(2026, 6, 11), "RENT", 2026)
        self.assertEqual((result.text_month, result.text_year), (None, 2026))
        self.assertEqual((result.period_month, result.period_year), (6, 2026))
        self.assertIn("ikut posting date", result.notes[-1])
        old = parse_period("Sewa JUL-DES 25", date(2026, 4, 20), "RENT", 2026)
        self.assertEqual(old.carry_over_year, 2025)


if __name__ == "__main__":
    unittest.main()

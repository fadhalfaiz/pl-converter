"""Pure rules for mapping GL text to a PL period."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import re


EXCLUDED_ACCOUNTS = frozenset(
    {"DIRECT LABOR", "SALARY AND COMPENSATION", "DEPRECIATION & AMORTIZATION"}
)

MONTH_ALIASES = {
    1: ("january", "januari", "jan"),
    2: ("february", "februari", "feb"),
    3: ("march", "maret", "mar"),
    4: ("april", "apr"),
    5: ("mei", "may"),
    6: ("june", "juni", "jun"),
    7: ("july", "juli", "jul"),
    8: ("agustus", "august", "agust", "agst", "agu", "ags", "agt", "aug"),
    9: ("september", "sept", "sep"),
    10: ("october", "oktober", "okt", "oct"),
    11: ("november", "nov"),
    12: ("december", "desember", "des", "dec"),
}
ALIAS_TO_MONTH = {alias: month for month, aliases in MONTH_ALIASES.items() for alias in aliases}
MONTH_PATTERN = re.compile(
    r"(?<![A-Za-z])(?:" + "|".join(sorted(ALIAS_TO_MONTH, key=len, reverse=True)) + r")(?![A-Za-z])",
    re.IGNORECASE,
)
FOUR_DIGIT_YEAR = re.compile(r"(?<!\d)(20\d{2})(?!\d)")
SHORT_YEAR_AFTER_MONTH = re.compile(r"^\s*[-/.'’]*\s*(\d{2})(?!\d)")
TRAILING_SHORT_YEAR = re.compile(r"(?:^|\s)(25|26)\s*[-–]?\s*$")


@dataclass(frozen=True)
class PeriodResult:
    text_month: int | None
    text_year: int | None
    period_month: int | None
    period_year: int | None
    carry_over_year: int | None
    notes: tuple[str, ...] = ()

    @property
    def period_label(self) -> str:
        if self.carry_over_year is not None:
            return f"CO {self.carry_over_year}"
        assert self.period_month is not None and self.period_year is not None
        return f"{self.period_month:02d}-{self.period_year}"


def clean_text(value: object) -> str:
    """Normalize Excel text without changing the underlying GL value."""
    if value is None:
        return ""
    return " ".join(str(value).replace("\t", " ").replace('"', "").replace("“", "").replace("”", "").split())


def _posting_date(value: date | datetime) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raise ValueError("Posting Date harus berupa tanggal Excel yang valid.")


def parse_period(
    text: object, posting_date: date | datetime, gl_account: object,
    report_year: int, ytd_month: int = 12,
) -> PeriodResult:
    """Return editable text month/year, effective PL period, and review notes."""
    posting = _posting_date(posting_date)
    if not 1900 <= report_year <= 9999:
        raise ValueError("Tahun Laporan tidak valid.")
    if not 1 <= ytd_month <= 12:
        raise ValueError("Bulan YTD harus antara Januari dan Desember.")
    account = str(gl_account or "").strip().upper()
    normalized = clean_text(text)
    matches = list(MONTH_PATTERN.finditer(normalized))
    month_match = matches[-1] if matches else None
    month = ALIAS_TO_MONTH[month_match.group().lower()] if month_match else None
    year = None
    inferred_year = False
    notes: list[str] = []

    if month_match:
        suffix = normalized[month_match.end():]
        short = SHORT_YEAR_AFTER_MONTH.match(suffix)
        if short:
            year = 2000 + int(short.group(1))
        else:
            four = FOUR_DIGIT_YEAR.search(suffix)
            if four:
                year = int(four.group(1))
            else:
                four_anywhere = FOUR_DIGIT_YEAR.search(normalized)
                if four_anywhere:
                    year = int(four_anywhere.group(1))
        if year is None:
            inferred_year = True
            year = report_year - 1 if month > ytd_month else report_year
            notes.append(f"Bulan tanpa tahun -> tahun dianggap {year}")
        if len(matches) > 1:
            notes.append(
                "Rentang bulan tahun laporan -> ikut posting date"
                if year == report_year else "Rentang bulan, dipakai bulan akhir"
            )
        if not inferred_year and year < posting.year and month <= posting.month:
            notes.append(
                f"Cek: tahun {year} tapi bulan <= bulan posting, kemungkinan salah ketik"
            )
    else:
        four = FOUR_DIGIT_YEAR.search(normalized)
        if four:
            year = int(four.group(1))
        else:
            trailing = TRAILING_SHORT_YEAR.search(normalized)
            if trailing and re.search(r"[A-Za-z]", normalized[:trailing.start()]):
                year = 2000 + int(trailing.group(1))
                notes.append("Tahun dibaca dari angka di akhir teks, mohon konfirmasi")

    if account in EXCLUDED_ACCOUNTS:
        return PeriodResult(month, year, posting.month, posting.year, None,
                            ("GL dikecualikan -> ikut posting date",))
    if year is not None and year < report_year:
        return PeriodResult(month, year, None, None, year, tuple(notes))
    if len(matches) > 1 and year == report_year:
        # Keep G editable. A blank F makes the H formula use Posting Date;
        # entering F manually remains a simple override in Excel or the app.
        return PeriodResult(None, year, posting.month, posting.year, None, tuple(notes))
    if month is not None:
        return PeriodResult(month, year, month, year, None, tuple(notes))
    return PeriodResult(None, year, posting.month, posting.year, None, tuple(notes))


def effective_period(
    posting_date: date | datetime, gl_account: object, report_year: int,
    text_month: int | None, text_year: int | None,
) -> tuple[int | None, int | None, int | None]:
    """Mirror the Excel H-column formula after a user edits F/G."""
    posting = _posting_date(posting_date)
    if str(gl_account or "").strip().upper() in EXCLUDED_ACCOUNTS:
        return posting.month, posting.year, None
    if text_year is not None and text_year < report_year:
        return None, None, text_year
    if text_year is not None and text_month is not None:
        return text_month, text_year, None
    return posting.month, posting.year, None

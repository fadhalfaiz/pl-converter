"""Streamlit interface for converting a single-project SAP GL into PL REALISASI."""

from __future__ import annotations

import hashlib
import json
from numbers import Real
from pathlib import Path
import re

import pandas as pd
import streamlit as st

from pl_builder import (
    apply_edits, build_workbook, default_justifications,
    preview_pl, read_gl, recalculate_with_libreoffice,
)


DEFAULT_TEMPLATE = Path(__file__).with_name("Tamplate_PL.xlsx")
DEFAULT_CREATOR_NAME = "Fadhlillah Rifqi H"
DEFAULT_CREATOR_ROLE = "UH Performance BBM"
DEFAULT_REVIEWER_NAME = "Rizky Farlyan"
DEFAULT_REVIEWER_ROLE = "Manager TOPM Dept."
DEFAULT_APPROVER_NAME = "Rahman Ilahi"
DEFAULT_APPROVER_ROLE = "GM Transportation Operation"
MONTH_NAMES_ID = (
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
)


def _number(value: object) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, Real) and not isinstance(value, bool):
        if value == 0:
            return "-"
        amount = f"{abs(value):,.0f}".replace(",", ".")
        return f"- {amount}" if value < 0 else amount
    return str(value)


def _percent(value: object) -> str:
    if pd.isna(value):
        return ""
    if value == 0:
        return "-"
    amount = f"{abs(value):.1%}".replace(".", ",")
    return f"- {amount}" if value < 0 else amount


def _safe_filename(project: str, year: int, ytd_month: int) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", project).strip("_")
    return f"PL_{name}_{year}_YTD_{ytd_month:02d}.xlsx"


def main() -> None:
    st.set_page_config(page_title="PL REALISASI", layout="wide")
    st.title("PL REALISASI")
    st.caption("Unggah GL SAP, periksa penentuan periode, lalu unduh laporan Excel berumus.")

    creator_name = st.text_input("Pembuat", value=DEFAULT_CREATOR_NAME)
    creator_role = st.text_input("Jabatan", value=DEFAULT_CREATOR_ROLE)
    reviewer_name = st.text_input("Mengetahui", value=DEFAULT_REVIEWER_NAME)
    reviewer_role = st.text_input("Jabatan Mengetahui", value=DEFAULT_REVIEWER_ROLE)
    approver_name = st.text_input("Menyetujui", value=DEFAULT_APPROVER_NAME)
    approver_role = st.text_input("Jabatan Menyetujui", value=DEFAULT_APPROVER_ROLE)
    gl_file = st.file_uploader("File GL SAP (.xlsx)", type="xlsx")
    year = st.number_input("Tahun Laporan", min_value=1900, max_value=9999, value=2026, step=1)
    ytd_month = st.selectbox(
        "Bulan YTD (sampai)", range(1, 13), index=11,
        format_func=lambda month: MONTH_NAMES_ID[month - 1],
    )
    st.caption("Tiga akun GL yang dikecualikan selalu mengikuti Posting Date. Untuk akun lain, bulan tanpa tahun yang melewati YTD dianggap tahun sebelumnya.")

    if st.button("Proses", type="primary"):
        st.session_state["pl_gl"] = None
        if gl_file is None:
            st.error("Unggah file GL terlebih dahulu.")
        else:
            try:
                template_bytes = DEFAULT_TEMPLATE.read_bytes()
                gl = read_gl(gl_file.getvalue(), int(year), ytd_month)
                preview_pl(gl, template_bytes, int(year), ytd_month)
                st.session_state["pl_gl"] = gl
                st.session_state["pl_template"] = template_bytes
                st.session_state["pl_year"] = int(year)
                st.session_state["pl_ytd_month"] = ytd_month
                st.session_state["pl_output"] = None
                st.session_state["pl_output_key"] = None
                st.session_state["pl_just_defaults"] = {}
                for key in list(st.session_state):
                    if key.startswith(("justification_", "review_editor_")):
                        del st.session_state[key]
            except (ValueError, OSError) as exc:
                st.error(str(exc))

    gl = st.session_state.get("pl_gl")
    if gl is None:
        return
    template_bytes = st.session_state["pl_template"]
    report_year = st.session_state["pl_year"]
    report_month = st.session_state["pl_ytd_month"]
    st.subheader(f"{gl.project} · YTD s.d. {MONTH_NAMES_ID[report_month-1]} {report_year}")
    st.caption(f"{len(gl.records):,} baris GL · Sheet sumber: {gl.sheet_name}")

    preview_tab, review_tab, carry_tab = st.tabs(
        ["Preview PL", "Perlu Dicek", "Ringkasan Carry Over"]
    )

    with review_tab:
        review = gl.records[gl.records["Catatan"].fillna("") != ""]
        review_posting = pd.to_datetime(review["Posting Date"])
        review = review.loc[
            review_posting.dt.year.eq(report_year) & review_posting.dt.month.le(report_month)
        ]
        note_types = sorted({part.strip() for note in review["Catatan"] for part in note.split(";")})
        selected_notes = st.multiselect("Filter jenis catatan", note_types)
        if selected_notes:
            review = review[review["Catatan"].apply(lambda note: any(tag in note for tag in selected_notes))]
        columns = ["Baris Excel", "GL Account", "Posting Date", "Text", "Bulan (Text)", "Tahun (Text)", "Periode PL", "Catatan"]
        edited = st.data_editor(
            review[columns].copy(), hide_index=True, use_container_width=True,
            disabled=[col for col in columns if col not in ("Bulan (Text)", "Tahun (Text)")],
            column_config={
                "Bulan (Text)": st.column_config.NumberColumn("Bulan (Text)", min_value=1, max_value=12, step=1),
                "Tahun (Text)": st.column_config.NumberColumn("Tahun (Text)", min_value=1900, max_value=9999, step=1),
            },
            key="review_editor_" + hashlib.sha1("|".join(selected_notes).encode()).hexdigest()[:10],
        )
        try:
            apply_edits(gl, edited, report_year)
        except ValueError as exc:
            st.error(str(exc))

    try:
        preview = preview_pl(gl, template_bytes, report_year, report_month)
    except ValueError as exc:
        st.error(str(exc))
        return
    with preview_tab:
        view = preview.drop(columns="Baris").copy()
        formatted = pd.DataFrame([
            {"ITEMS": row["ITEMS"], **{
                col: (_percent(row[col]) if row["ITEMS"] == "GPM" else _number(row[col]))
                for col in view.columns[1:]
            }} for _, row in view.iterrows()
        ])
        st.dataframe(formatted, hide_index=True, use_container_width=True, height=560)
        st.caption("OPERATING REVENUE dibalik tanda dari GL; akun lain memakai tanda aslinya. TOTAL YTD tidak memasukkan Carry Over. Semua total dibatasi Posting Date sampai akhir bulan YTD.")

    with carry_tab:
        carry = gl.records[gl.records["Periode PL"] == f"CO {report_year-1}"].copy()
        carry_posting = pd.to_datetime(carry["Posting Date"])
        carry = carry.loc[
            carry_posting.dt.year.eq(report_year) & carry_posting.dt.month.le(report_month)
        ]
        if carry.empty:
            st.info("Tidak ada baris Carry Over untuk tahun sebelumnya.")
        else:
            carry["Nilai PL"] = carry["In profit center local currency"] * carry["GL Account"].eq("OPERATING REVENUE").map({True: -1, False: 1})
            summary = carry.groupby("GL Account", as_index=False)["Nilai PL"].sum()
            st.dataframe(summary.style.format({"Nilai PL": _number}), hide_index=True, use_container_width=True)
            account = st.selectbox("Lihat baris penyusun", summary["GL Account"].tolist())
            details = carry[carry["GL Account"] == account][
                ["Baris Excel", "Posting Date", "Text", "In profit center local currency", "Nilai PL", "Catatan"]
            ]
            st.dataframe(details.style.format({"In profit center local currency": _number, "Nilai PL": _number}),
                         hide_index=True, use_container_width=True)

    st.subheader("Justifikasi")
    defaults = default_justifications(preview, report_year)
    carry_rows = preview.loc[
        preview["Baris"].isin((4, 7, 8, 9, *range(11, 33)))
        & preview[f"Carry Over {report_year-1}"].ne(0), "Baris"
    ].tolist()
    old_defaults = st.session_state.setdefault("pl_just_defaults", {})
    justifications = []
    for index, (row, default) in enumerate(zip(carry_rows, defaults)):
        key = f"justification_{row}"
        if key not in st.session_state or st.session_state[key] == old_defaults.get(row):
            st.session_state[key] = default
        old_defaults[row] = default
        line = st.text_input(f"Poin {index + 1}", key=key)
        justifications.append(line)

    evaluation_rate = preview.loc[
        preview["Baris"].eq(48), f"TOTAL YTD {report_year}"
    ].iloc[0]
    operating_revenue = preview.loc[preview["Baris"].eq(4), f"TOTAL YTD {report_year}"].iloc[0]
    evaluation_rate = evaluation_rate / operating_revenue if operating_revenue else 0
    evaluation_percent = f"{abs(evaluation_rate):.2%}".replace(".", ",")
    if evaluation_rate < 0:
        evaluation_percent = "− " + evaluation_percent
    st.subheader("Hasil Evaluasi")
    st.write(
        "Nilai NPAT setelah dilakukan cleansing sebesar "
        + f"**{evaluation_percent}**"
    )

    edit_signature = gl.records[["Baris Excel", "Bulan (Text)", "Tahun (Text)"]].to_json(orient="records")
    fingerprint = hashlib.sha256(
        gl.source_bytes + template_bytes + str(report_year).encode() + str(report_month).encode()
        + edit_signature.encode() + json.dumps({
            "justifications": justifications,
            "creator_name": creator_name,
            "creator_role": creator_role,
            "reviewer_name": reviewer_name,
            "reviewer_role": reviewer_role,
            "approver_name": approver_name,
            "approver_role": approver_role,
        }, ensure_ascii=False).encode()
    ).hexdigest()
    if st.session_state.get("pl_output_key") != fingerprint:
        with st.spinner("Menyusun Excel dan memeriksa rumus..."):
            try:
                result = build_workbook(
                    gl, template_bytes, report_year, justifications,
                    ytd_month=report_month,
                    creator_name=creator_name,
                    creator_role=creator_role,
                    reviewer_name=reviewer_name,
                    reviewer_role=reviewer_role,
                    approver_name=approver_name,
                    approver_role=approver_role,
                )
                result, recalc_message = recalculate_with_libreoffice(result)
                st.session_state["pl_output"] = result
                st.session_state["pl_output_message"] = recalc_message
                st.session_state["pl_output_key"] = fingerprint
            except (ValueError, RuntimeError) as exc:
                st.error(str(exc))
                return
    st.caption(st.session_state["pl_output_message"])
    st.download_button(
        "Download Excel", data=st.session_state["pl_output"],
        file_name=_safe_filename(gl.project, report_year, report_month),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    main()

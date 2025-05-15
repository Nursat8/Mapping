#!/usr/bin/env python3
"""Streamlit app to map ESG / PAI / Taxonomy IDs into EquityRef_ESGdataMap.

📅 v2025‑05‑15‑c  ➜  *Keeps every literal “NA” or “N/A” that existed in the file.*
                 A cell is overwritten **only if**
                   1) its content is a placeholder (blank / NA / N/A), **and**
                   2) a matching ID is found in the relevant mapping file.

If no match is found, the original placeholder text remains untouched.
"""

from __future__ import annotations
import streamlit as st
import pandas as pd
from io import BytesIO

# The strings we regard as placeholders (case‑sensitive because that’s
# what appears in your sheets).
PLACEHOLDERS = {"", "NA", "N/A"}

# ────────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ────────────────────────────────────────────────────────────────────────────────

def read_id_series(path_or_buf, col_name: str, *, header: int | None = 0) -> pd.Series:
    """Return a Series of *Int64* IDs from *col_name* in an Excel file."""
    df = pd.read_excel(path_or_buf, header=header, dtype=str, keep_default_na=False)
    df.columns = df.columns.str.strip()
    if col_name not in df.columns:
        raise ValueError(f"Column '{col_name}' not found in {getattr(path_or_buf,'name',path_or_buf)}")
    ser = pd.to_numeric(df[col_name].str.strip(), errors="coerce").dropna().astype("Int64")
    return ser


def to_int(series: pd.Series) -> pd.Series:
    """Convert a string/object Series to nullable Int64 (NaNs where non‑numeric)."""
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def is_placeholder(series: pd.Series) -> pd.Series:
    """Boolean mask – True where the cell is considered a placeholder."""
    return series.isna() | series.astype(str).str.strip().isin(PLACEHOLDERS)


# ────────────────────────────────────────────────────────────────────────────────
# Streamlit app
# ────────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper – NA‑aware")

    # File upload widgets
    equity_file = st.file_uploader("Main file: EquityRef_ESGdataMap", type=["xlsx", "xls"])
    taxonomy_files = st.file_uploader(
        "Taxonomy mapping files (both FMP & NFC)", type=["xlsx", "xls"], accept_multiple_files=True
    )
    pai_file = st.file_uploader("PAI mapping: SFDR_Corporate_CPUPU_LYr_202503", type=["xlsx", "xls"])
    esg_file = st.file_uploader("ESG mapping: 111 (SP_ENTITY_ID)", type=["xlsx", "xls"])

    if st.button("Process & Fill Columns"):
        # Basic validation
        if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
            st.error("⬆️  Please upload *all* required files first.")
            return

        # ── Read main file (do *not* treat 'NA'/'N/A' as NaN) ──────────────
        df = pd.read_excel(equity_file, dtype=str, keep_default_na=False)

        id_col, tax_col, pai_col, esg_col = "Ids", "Taxonomy", "PAI", "ESG"
        for col in (id_col, tax_col, pai_col, esg_col):
            if col not in df.columns:
                df[col] = "NA"  # ensure column exists with placeholder text

        # Build numeric version of IDs for fast membership tests
        df["_id_num"] = to_int(df[id_col])

        # ── Gather mapping IDs ────────────────────────────────────────────
        taxonomy_ids: set[int] = set()
        for f in taxonomy_files:
            try:
                taxonomy_ids.update(read_id_series(f, "MI Key"))
            except Exception as e:
                st.warning(str(e))

        try:
            pai_ids: set[int] = set(read_id_series(pai_file, "KeyInstn"))
        except Exception as e:
            st.warning(str(e))
            pai_ids = set()

        try:
            esg_ids: set[int] = set(read_id_series(esg_file, "SP_ENTITY_ID", header=4))
        except Exception as e:
            st.warning(str(e))
            esg_ids = set()

        # ── Apply mappings (only overwrite placeholders) ───────────────────
        def fill_column(target_col: str, id_set: set[int]):
            mask = df["_id_num"].isin(id_set) & is_placeholder(df[target_col])
            df.loc[mask, target_col] = df.loc[mask, id_col]
            return int(mask.sum())

        n_tax = fill_column(tax_col, taxonomy_ids)
        n_pai = fill_column(pai_col, pai_ids)
        n_esg = fill_column(esg_col, esg_ids)

        # ── Report ────────────────────────────────────────────────────────
        st.success("Processing complete.")
        st.write({"Rows": len(df), "Taxonomy updated": n_tax, "PAI updated": n_pai, "ESG updated": n_esg})

        # Clean‑up helper column & export
        df.drop(columns=["_id_num"], inplace=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        st.download_button(
            label="Download Filled Excel",
            data=output,
            file_name="EquityRef_ESGdataMap_filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()

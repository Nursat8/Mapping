#!/usr/bin/env python3
"""Streamlit app to map ESG / PAI / Taxonomy IDs.

v2025‑05‑15‑b – Keeps literal "NA" / "N/A" strings.
Only replaces cells *currently* placeholder (blank, NaN, "NA", "N/A")
when a matching ID is found. Other existing values remain untouched.
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# Placeholder strings the user wants preserved unless overwritten
PLACEHOLDERS = {"", "NA", "N/A", "na", "n/a", "Na"}

# ────────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ────────────────────────────────────────────────────────────────────────────────

def read_id_column(file, col_name: str, header: int | None = 0):
    """Read *col_name* from *file* and return an Int64 Series of IDs."""
    df = pd.read_excel(file, header=header, dtype=str)
    df.columns = df.columns.str.strip()
    if col_name not in df.columns:
        raise ValueError(f"Column '{col_name}' not found in {getattr(file,'name',file)}")
    series = pd.to_numeric(df[col_name], errors="coerce").dropna().astype("Int64")
    return series


def numeric_ids(series: pd.Series) -> pd.Series:
    """Convert a Series to Int64, coercing errors to NaNs (keeps original object dtype)."""
    return pd.to_numeric(series, errors="coerce").astype("Int64")


def placeholder_mask(series: pd.Series) -> pd.Series:
    """True where the cell is NaN, blank, or one of the placeholder strings."""
    return (
        series.isna()
        | series.astype(str).str.strip().isin(PLACEHOLDERS)
    )


# ────────────────────────────────────────────────────────────────────────────────
# Streamlit app
# ────────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper (keeps NA placeholders)")

    equity_file = st.file_uploader("Main file: EquityRef_ESGdataMap", type=["xlsx", "xls"])
    taxonomy_files = st.file_uploader(
        "Taxonomy mapping files (FMP & NFC)", type=["xlsx", "xls"], accept_multiple_files=True
    )
    pai_file = st.file_uploader("PAI mapping (SFDR_Corporate_CPUPU_LYr_202503)", type=["xlsx", "xls"])
    esg_file = st.file_uploader("ESG mapping (111 – SP_ENTITY_ID)", type=["xlsx", "xls"])

    if st.button("Process & Fill Columns"):
        if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
            st.error("Please upload all required files.")
            return

        # Read main file WITHOUT auto‑converting 'NA' to NaN
        df_main = pd.read_excel(equity_file, dtype=str)
        id_col, tax_col, pai_col, esg_col = "Ids", "Taxonomy", "PAI", "ESG"

        if id_col not in df_main.columns:
            st.error(f"Column '{id_col}' not found in main file.")
            return

        # Convert Ids to Int64 for matching but keep original str dtype for export
        df_main[id_col + "_num"] = numeric_ids(df_main[id_col])

        # Ensure target columns exist
        for col in (tax_col, pai_col, esg_col):
            if col not in df_main.columns:
                df_main[col] = "NA"  # default placeholder

        # ── Taxonomy mapping ────────────────────────────────────────────────
        taxonomy_ids_all = []
        for f in taxonomy_files:
            try:
                taxonomy_ids_all.append(read_id_column(f, "MI Key"))
            except Exception as e:
                st.warning(str(e))
        taxonomy_set = set(pd.concat(taxonomy_ids_all).unique()) if taxonomy_ids_all else set()
        mask_tax = (
            df_main[id_col + "_num"].isin(taxonomy_set)
            & placeholder_mask(df_main[tax_col])
        )
        df_main.loc[mask_tax, tax_col] = df_main.loc[mask_tax, id_col]

        # ── PAI mapping ────────────────────────────────────────────────────
        try:
            pai_set = set(read_id_column(pai_file, "KeyInstn"))
        except Exception as e:
            st.warning(str(e))
            pai_set = set()
        mask_pai = (
            df_main[id_col + "_num"].isin(pai_set)
            & placeholder_mask(df_main[pai_col])
        )
        df_main.loc[mask_pai, pai_col] = df_main.loc[mask_pai, id_col]

        # ── ESG mapping ────────────────────────────────────────────────────
        try:
            esg_set = set(read_id_column(esg_file, "SP_ENTITY_ID", header=4))
        except Exception as e:
            st.warning(str(e))
            esg_set = set()
        mask_esg = (
            df_main[id_col + "_num"].isin(esg_set)
            & placeholder_mask(df_main[esg_col])
        )
        df_main.loc[mask_esg, esg_col] = df_main.loc[mask_esg, id_col]

        # Diagnostics
        st.success("Finished processing!")
        st.write(
            {
                "Rows": len(df_main),
                "Taxonomy replaced": int(mask_tax.sum()),
                "PAI replaced": int(mask_pai.sum()),
                "ESG replaced": int(mask_esg.sum()),
            }
        )

        # Clean‑up: drop helper numeric column
        df_main.drop(columns=[id_col + "_num"], inplace=True)

        # Export
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_main.to_excel(writer, index=False)
        output.seek(0)
        st.download_button(
            "Download Filled Excel",
            output,
            "EquityRef_ESGdataMap_filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()

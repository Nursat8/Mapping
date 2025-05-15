#!/usr/bin/env python3
"""Streamlit app to map ESG / PAI / Taxonomy IDs.
   v2025‑05‑15 – Added logic to treat the literal strings
   "NA", "N/A", and "na" as missing and ONLY overwrite empty/NA cells.
"""

import streamlit as st
import pandas as pd
from io import BytesIO

NA_STRINGS = ["NA", "N/A", "na", "Na", "n/a"]

# ────────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ────────────────────────────────────────────────────────────────────────────────

def read_id_column(file, col_name: str, header: int | None = 0):
    """Read a single ID column as Int64 Series, stripping NA‑like strings."""
    df = pd.read_excel(file, header=header, dtype={col_name: str})
    df.columns = df.columns.str.strip()
    if col_name not in df.columns:
        raise ValueError(f"Column '{col_name}' not found in {getattr(file,'name',file)}")
    ser = df[col_name].replace(NA_STRINGS, pd.NA).dropna()
    return pd.to_numeric(ser, errors="coerce").dropna().astype("Int64")


def ids_series(df: pd.DataFrame, column: str) -> pd.Series:
    """Return df[column] coerced to Int64, treating NA strings as missing."""
    ser = df[column].replace(NA_STRINGS, pd.NA)
    return pd.to_numeric(ser, errors="coerce").astype("Int64")


# ────────────────────────────────────────────────────────────────────────────────
# Streamlit app
# ────────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper (v2025‑05‑15)")

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

        # Read main file – treat NA-like strings as NaNs right away
        df_main = pd.read_excel(equity_file, dtype=str, na_values=NA_STRINGS)
        id_col, tax_col, pai_col, esg_col = "Ids", "Taxonomy", "PAI", "ESG"

        if id_col not in df_main.columns:
            st.error(f"Column '{id_col}' not found in main file.")
            return

        df_main[id_col] = ids_series(df_main, id_col)

        # Make sure target columns exist & normalised
        for col in (tax_col, pai_col, esg_col):
            if col not in df_main.columns:
                df_main[col] = pd.Series([pd.NA] * len(df_main), dtype="Int64")
            else:
                df_main[col] = ids_series(df_main, col)

        # ── Taxonomy mapping ────────────────────────────────────────────────
        taxonomy_ids_all = []
        for f in taxonomy_files:
            try:
                taxonomy_ids_all.append(read_id_column(f, "MI Key"))
            except Exception as e:
                st.warning(str(e))
        taxonomy_ids = set(pd.concat(taxonomy_ids_all).unique()) if taxonomy_ids_all else set()
        mask_tax = df_main[id_col].isin(taxonomy_ids) & df_main[tax_col].isna()
        df_main.loc[mask_tax, tax_col] = df_main.loc[mask_tax, id_col]

        # ── PAI mapping ────────────────────────────────────────────────────
        try:
            pai_ids = set(read_id_column(pai_file, "KeyInstn"))
            mask_pai = df_main[id_col].isin(pai_ids) & df_main[pai_col].isna()
            df_main.loc[mask_pai, pai_col] = df_main.loc[mask_pai, id_col]
        except Exception as e:
            st.warning(str(e))
            mask_pai = pd.Series(False, index=df_main.index)

        # ── ESG mapping ────────────────────────────────────────────────────
        try:
            esg_ids = set(read_id_column(esg_file, "SP_ENTITY_ID", header=4))
            mask_esg = df_main[id_col].isin(esg_ids) & df_main[esg_col].isna()
            df_main.loc[mask_esg, esg_col] = df_main.loc[mask_esg, id_col]
        except Exception as e:
            st.warning(str(e))
            mask_esg = pd.Series(False, index=df_main.index)

        # Diagnostics
        st.success("Finished processing!")
        st.write(
            {
                "Rows": len(df_main),
                "Taxonomy filled": int(mask_tax.sum()),
                "PAI filled": int(mask_pai.sum()),
                "ESG filled": int(mask_esg.sum()),
            }
        )

        # Export to Excel
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

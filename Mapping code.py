#!/usr/bin/env python3
"""Streamlit app to map ESG / PAI / Taxonomy IDs.

Key fixes 2025‑05‑14:
• Convert all IDs to pandas Int64 before matching → no more "12345.0" vs "12345" problems.
• Explicit header row for ESG file (row 5, i.e. header=4).
• Extra diagnostics: after processing, show how many matches were found for each column.
• Defensive column look‑ups with .str.strip() to avoid trailing spaces in headers.
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# ────────────────────────────────────────────────────────────────────────────────
# Helper functions
# ────────────────────────────────────────────────────────────────────────────────

def read_id_column(file, col_name: str, header: int | None = 0):
    """Return a Series of Int64 IDs from `col_name` in the given Excel file."""
    try:
        df = pd.read_excel(file, header=header, usecols=[col_name])
    except ValueError:
        # Try stripping spaces in the column names as a fallback
        df = pd.read_excel(file, header=header)
        df.columns = df.columns.str.strip()
        if col_name.strip() not in df.columns:
            raise
        df = df[[col_name.strip()]]
    ids = pd.to_numeric(df[col_name].squeeze().rename("id"), errors="coerce")
    return ids.dropna().astype("Int64")


def ids_series(df: pd.DataFrame, column: str) -> pd.Series:
    """Return Int64 Series from df[column], coercing errors to NaN."""
    return pd.to_numeric(df[column], errors="coerce").astype("Int64")


# ────────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ────────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper (v2025‑05‑14)")

    equity_file = st.file_uploader("EquityRef_ESGdataMap (main file)", type=["xlsx", "xls"])
    taxonomy_files = st.file_uploader(
        "EUTaxRevShare Taxonomy mapping files (FMP & NFC)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
    pai_file = st.file_uploader("SFDR_Corporate_CPUPU_LYr_202503 (PAI)", type=["xlsx", "xls"])
    esg_file = st.file_uploader("SPGlobal Export / 111 file (ESG)", type=["xlsx", "xls"])

    if st.button("Process & Fill Columns"):
        if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
            st.error("⬆️ Please upload *all* required files first.")
            return

        # ── Read main file ────────────────────────────────────────────────────
        df_main = pd.read_excel(equity_file)
        id_col, tax_col, pai_col, esg_col = "Ids", "Taxonomy", "PAI", "ESG"

        if id_col not in df_main.columns:
            st.error(f"❌ Column ‘{id_col}’ not found in main file.")
            return

        # Convert Ids to Int64 for robust matching
        df_main[id_col] = ids_series(df_main, id_col)

        # Create empty target cols if missing
        for col in (tax_col, pai_col, esg_col):
            if col not in df_main.columns:
                df_main[col] = pd.Series([pd.NA] * len(df_main), dtype="Int64")
            else:
                df_main[col] = ids_series(df_main, col)  # normalise existing data

        # ── Taxonomy (MI Key) ────────────────────────────────────────────────
        taxonomy_ids = pd.Series(dtype="Int64")
        for f in taxonomy_files:
            try:
                ids = read_id_column(f, "MI Key")
                taxonomy_ids = pd.concat([taxonomy_ids, ids])
            except Exception as e:
                st.warning(f"Skipping {getattr(f, 'name', f)} – couldn’t read MI Key: {e}")
        taxonomy_set = set(taxonomy_ids.unique())
        mask_tax = df_main[id_col].isin(taxonomy_set)
        df_main.loc[mask_tax, tax_col] = df_main.loc[mask_tax, id_col]

        # ── PAI (KeyInstn) ───────────────────────────────────────────────────
        try:
            pai_ids = set(read_id_column(pai_file, "KeyInstn").unique())
            mask_pai = df_main[id_col].isin(pai_ids)
            df_main.loc[mask_pai, pai_col] = df_main.loc[mask_pai, id_col]
        except Exception as e:
            st.warning(f"PAI file issue: {e}")

        # ── ESG (SP_ENTITY_ID) – header row at 5th row (header=4) ────────────
        try:
            esg_ids = set(read_id_column(esg_file, "SP_ENTITY_ID", header=4).unique())
            mask_esg = df_main[id_col].isin(esg_ids)
            df_main.loc[mask_esg, esg_col] = df_main.loc[mask_esg, id_col]
        except Exception as e:
            st.warning(f"ESG file issue: {e}")

        # ── Diagnostics ──────────────────────────────────────────────────────
        st.success("Finished processing!")
        st.write(
            {
                "Total rows": len(df_main),
                "Taxonomy matches": int(mask_tax.sum()),
                "PAI matches": int(mask_pai.sum()) if "mask_pai" in locals() else 0,
                "ESG matches": int(mask_esg.sum()) if "mask_esg" in locals() else 0,
            }
        )

        # ── Export ───────────────────────────────────────────────────────────
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_main.to_excel(writer, index=False)
        output.seek(0)
        st.download_button(
            "Download Filled Excel", output, "EquityRef_ESGdataMap_filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()

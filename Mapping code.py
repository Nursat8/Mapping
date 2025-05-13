#!/usr/bin/env python3
import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper")

    # 1) Upload input files
    equity_file = st.file_uploader("Upload EquityRef_ESGdataMap (main file)", type=["xlsx", "xls"])
    taxonomy_files = st.file_uploader("Upload Taxonomy mapping files (2x)", type=["xlsx", "xls"], accept_multiple_files=True)
    pai_file = st.file_uploader("Upload PAI mapping file (SFDR_Corporate_CPUPU_LYr_202503)", type=["xlsx", "xls"])
    esg_file = st.file_uploader("Upload ESG mapping file (SPGlobal_Export_4-11-2025 file)", type=["xlsx", "xls"])

    if st.button("Process & Fill Columns"):
        if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
            st.error("Please upload: main file, both taxonomy files, PAI file, and ESG file.")
            return

        # Read main DataFrame
        df_main = pd.read_excel(equity_file)
        id_col = 'Ids'       # column G
        taxonomy_col = 'Taxonomy'
        pai_col = 'PAI'
        esg_col = 'ESG'

        # Validate ID column in main
        if id_col not in df_main.columns:
            st.error(f"Column '{id_col}' not found in main file.")
            return

        # Initialize output columns if missing
        for col in [taxonomy_col, pai_col, esg_col]:
            if col not in df_main.columns:
                df_main[col] = ""

        # --- Taxonomy mapping ---
        taxonomy_ids = set()
        for f in taxonomy_files:
            try:
                # Read only the MI Key column
                df_tax = pd.read_excel(f, usecols=['MI Key'])
                taxonomy_ids.update(df_tax['MI Key'].dropna().astype(str))
            except Exception as e:
                st.warning(f"Error reading taxonomy file {getattr(f, 'name', '')}: {e}")

        # Mark matches in Taxonomy column
        df_main[taxonomy_col] = df_main[id_col].astype(str).where(
            df_main[id_col].astype(str).isin(taxonomy_ids),
            ""
        )

        # --- PAI mapping ---
        try:
            df_pai = pd.read_excel(pai_file, usecols=['KeyInstn'])
            pai_ids = set(df_pai['KeyInstn'].dropna().astype(str))
            df_main[pai_col] = df_main[id_col].astype(str).where(
                df_main[id_col].astype(str).isin(pai_ids),
                ""
            )
        except Exception as e:
            st.warning(f"Error reading PAI file: {e}")

        # --- ESG mapping ---
        try:
            # Header row is in row 5 (0-indexed 4), data from row 8 onward
            df_esg = pd.read_excel(esg_file, header=4, usecols=['SP_ENTITY_ID'])
            esg_ids = set(df_esg['SP_ENTITY_ID'].dropna().astype(str))
            df_main[esg_col] = df_main[id_col].astype(str).where(
                df_main[id_col].astype(str).isin(esg_ids),
                ""
            )
        except Exception as e:
            st.warning(f"Error reading ESG file: {e}")

        # Export to Excel for download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_main.to_excel(writer, index=False)
        output.seek(0)

        st.download_button(
            label="Download Filled Excel",
            data=output,
            file_name="EquityRef_ESGdataMap_filled.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()

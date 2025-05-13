#!/usr/bin/env python3
import streamlit as st
import pandas as pd
from io import BytesIO


def main():
    st.set_page_config(page_title="EquityRef ESG Data Mapper", layout="wide")
    st.title("EquityRef ESG Data Mapper")

    # 1) Upload input files
    equity_file = st.file_uploader(
        "Upload EquityRef_ESGdataMap (main file)", type=["xlsx", "xls"]
    )
    taxonomy_files = st.file_uploader(
        "Upload Taxonomy mapping files: EUTaxRevShare_As_Reported_FMP_AllYr_20241220 and NFC (2 files)",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
    pai_file = st.file_uploader(
        "Upload PAI mapping file (SFDR_Corporate_CPUPU_LYr_202503)",
        type=["xlsx", "xls"]
    )
    esg_file = st.file_uploader(
        "Upload ESG mapping file (111 with SP_ENTITY_ID)",
        type=["xlsx", "xls"]
    )

    if st.button("Process & Fill Columns"):
        # Validate uploads
        if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
            st.error(
                "Please upload: main file, both taxonomy files, PAI file, and ESG file."
            )
            return

        # Read main sheet
        df_main = pd.read_excel(equity_file)
        id_col = 'Ids'       # column G in EquityRef_ESGdataMap
        taxonomy_col = 'Taxonomy'  # column D
        pai_col = 'PAI'           # column E
        esg_col = 'ESG'           # column F

        if id_col not in df_main.columns:
            st.error(f"'{id_col}' not found in main file. Please check column G header.")
            return

        # Initialize target columns
        for col in [taxonomy_col, pai_col, esg_col]:
            if col not in df_main.columns:
                df_main[col] = ""

        # --- Taxonomy: check MI Key (D1) in two mapping files ---
        taxonomy_ids = set()
        for f in taxonomy_files:
            try:
                df_tax = pd.read_excel(f, usecols=['MI Key'])
                taxonomy_ids.update(
                    df_tax['MI Key'].dropna().astype(str).tolist()
                )
            except Exception as e:
                st.warning(
                    f"Could not read taxonomy file {getattr(f, 'name', '')}: {e}"
                )
        df_main[taxonomy_col] = df_main[id_col].astype(str).where(
            df_main[id_col].astype(str).isin(taxonomy_ids), ""
        )

        # --- PAI: check KeyInstn (E1) in SFDR file ---
        try:
            df_pai = pd.read_excel(pai_file, usecols=['KeyInstn'])
            pai_ids = set(df_pai['KeyInstn'].dropna().astype(str).tolist())
            df_main[pai_col] = df_main[id_col].astype(str).where(
                df_main[id_col].astype(str).isin(pai_ids), ""
            )
        except Exception as e:
            st.warning(f"Could not read PAI file: {e}")

        # --- ESG: check SP_ENTITY_ID (B1) in 111 file ---
        try:
            df_esg = pd.read_excel(esg_file, usecols=['SP_ENTITY_ID'])
            esg_ids = set(df_esg['SP_ENTITY_ID'].dropna().astype(str).tolist())
            df_main[esg_col] = df_main[id_col].astype(str).where(
                df_main[id_col].astype(str).isin(esg_ids), ""
            )
        except Exception as e:
            st.warning(f"Could not read ESG file: {e}")

        # --- Export result ---
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

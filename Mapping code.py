import streamlit as st
import pandas as pd
from io import BytesIO

st.title("EquityRef ESG Data Mapper")

# Upload inputs
equity_file = st.file_uploader("Upload EquityRef_ESGdataMap file", type=["xlsx", "xls"])  
taxonomy_files = st.file_uploader(
    "Upload Taxonomy mapping files (FMP & NFC)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)
pai_file = st.file_uploader(
    "Upload PAI mapping file (SFDR_Corporate_CPUPU_LYr_202503)", 
    type=["xlsx", "xls"]
)
esg_file = st.file_uploader(
    "Upload ESG mapping file (SPGlobal_Export_4-11-2025_97777e98-3cf6-4636-9833-ceca80a4f6c6)", 
    type=["xlsx", "xls"]
)

if st.button("Process & Fill Columns"):
    # Check all files provided
    if not equity_file or len(taxonomy_files) < 2 or not pai_file or not esg_file:
        st.error("Please upload the main file, both taxonomy files, the PAI file, and the ESG file.")
    else:
        # Read main DataFrame
        df_main = pd.read_excel(equity_file)
        
        # Define column names
        id_col = 'Ids'
        taxonomy_col = 'Taxonomy'
        pai_col = 'PAI'
        esg_col = 'ESG'
        
        # Ensure ID column exists
        if id_col not in df_main.columns:
            st.error(f"Column '{id_col}' not found in main file.")
        else:
            # Read mapping DataFrames
            tax_dfs = []
            for f in taxonomy_files:
                try:
                    tax_dfs.append(pd.read_excel(f))
                except Exception as e:
                    st.error(f"Failed to read taxonomy file: {e}")

            df_pai_map = pd.read_excel(pai_file)
            df_esg_map = pd.read_excel(esg_file)

            # Initialize output columns if missing
            for col in [taxonomy_col, pai_col, esg_col]:
                if col not in df_main.columns:
                    df_main[col] = ''

            # Fill Taxonomy
            mask_tax = pd.Series(False, index=df_main.index)
            for tax_df in tax_dfs:
                if id_col in tax_df.columns:
                    mask_tax = mask_tax | df_main[id_col].isin(tax_df[id_col])
                else:
                    st.warning(f"Column '{id_col}' not found in one of the taxonomy mapping files.")
            df_main.loc[mask_tax, taxonomy_col] = df_main.loc[mask_tax, id_col]

            # Fill PAI
            if id_col in df_pai_map.columns:
                mask_pai = df_main[id_col].isin(df_pai_map[id_col])
                df_main.loc[mask_pai, pai_col] = df_main.loc[mask_pai, id_col]
            else:
                st.warning(f"Column '{id_col}' not found in PAI mapping file.")

            # Fill ESG
            if id_col in df_esg_map.columns:
                mask_esg = df_main[id_col].isin(df_esg_map[id_col])
                df_main.loc[mask_esg, esg_col] = df_main.loc[mask_esg, id_col]
            else:
                st.warning(f"Column '{id_col}' not found in ESG mapping file.")

            # Prepare for download
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

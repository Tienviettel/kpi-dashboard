import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import os
import mplcursors

st.set_page_config(page_title="KPI Evaluation Dashboard", layout="wide")
st.title("KPI Evaluation Dashboard")

# --- Sidebar: Upload input files and formula file ---
st.sidebar.header("1. Upload Files")
input_files = st.sidebar.file_uploader(
    "Upload Performance Management Excel files (multiple)",
    type=["xlsx"], accept_multiple_files=True
)
formula_file = st.sidebar.file_uploader(
    "Upload KPI_formula.xlsx (sheet 'KPI')", type=["xlsx"]
)

# === Utility functions from existing code ===
def process_excel(file_buffer, is_tdd=False):
    xl = pd.ExcelFile(file_buffer)
    sheet_names = [s for s in xl.sheet_names if s.strip().lower() != "kpi(counter)"]
    if not sheet_names:
        return pd.DataFrame()
    df = xl.parse(sheet_names[0])
    # clean column names
    df.columns = [col.replace("_FDD","").replace("_TDD","") for col in df.columns]
    if is_tdd:
        replace_dict_tdd = {
            "PRB Number Used on Downlink Channel": "LTE DL Physical Resource Block_Used",
            "PRB Number Available on Downlink Channel": "LTE DL Physical Resource Block_Available",
            "PRB Number Used on Uplink Channel": "LTE UL Physical Resource Block_Used",
            "PRB Number Available on Uplink Channel": "LTE UL Physical Resource Block_Available",
            "LTE Upload User Throughput_denumerator": "LTE Upload User Throughput_denumerato",
            "UL padding denumerator": "UL padding denominator"
        }
        df.rename(columns=replace_dict_tdd, inplace=True)
    return df

# Load KPI formulas mapping
@st.cache_data
def load_formulas(file_buffer):
    book = pd.ExcelFile(file_buffer)
    dfk = book.parse("KPI")
    dfk.columns = [str(c).strip() for c in dfk.columns]
    # build kpi_config from sheet
    kpi_config = {}
    formulas = {}
    for col in dfk.columns[1:]:
        kpi_name = col.strip()
        formula = dfk[col].dropna().iloc[0] if not dfk[col].dropna().empty else '((A-B)/B)*100'
        formulas[kpi_name] = str(formula).strip()
    return formulas

# Compute KPI values per row
def compute_kpis(df_raw, kpi_formulas, kpi_columns):
    # apply each formula on summed raw series
    df_kpi = pd.DataFrame()
    for kpi, formula in kpi_formulas.items():
        # placeholder: use eval on A and B
        # In table step, sum raw columns for before/after; in chart step, group by day
        pass
    return df_kpi

if input_files and formula_file:
    # --- Load and concatenate input data ---
    dfs = []
    for f in input_files:
        is_tdd = "TDD" in f.name.upper()
        dfs.append(process_excel(f, is_tdd))
    df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

    st.sidebar.success("Files loaded.")

    # --- Preprocess raw dataframe ---
    df['Begin Time'] = pd.to_datetime(df.get('Begin Time', pd.NaT), errors='coerce')
    df['Day'] = df['Begin Time'].dt.strftime('%d/%m/%Y')
    # remove commas and convert
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = (df[col].str.replace(',','')
                         .replace(['','nan','NaN','None'],'0'))
        df[col] = pd.to_numeric(df[col], errors='ignore')

    # --- Load formulas definitions ---
    raw_formulas = load_formulas(formula_file)
    # build kpi_columns as needed (user should customize mapping)
    # Example mapping: kpi_columns = { 'ePS-CSSR': ['LTE RRC Setup Attempt', 'LTE RRC Setup Success', ...], ... }
    kpi_columns = {}  # TODO: fill based on original KPI_Evaluation.py

    # --- Sidebar: Date selection ---
    st.sidebar.header("2. Select Dates")
    days = sorted(df['Day'].dropna().unique(), key=lambda d: datetime.strptime(d,'%d/%m/%Y'))
    before = st.sidebar.multiselect("Before Action Dates", days)
    after = st.sidebar.multiselect("After Action Dates", days)

    # --- Calculate & display KPI table ---
    if st.sidebar.button("Calculate Table"):
        if not before or not after:
            st.warning("Please select both Before and After dates.")
        else:
            # sum raw columns
            df_before = df[df['Day'].isin(before)]
            df_after  = df[df['Day'].isin(after)]
            summed_before = df_before.sum(numeric_only=True)
            summed_after  = df_after.sum(numeric_only=True)
            # calculate each KPI
            table = []
            for kpi, formula in raw_formulas.items():
                A = summed_after.get(kpi, 0)
                B = summed_before.get(kpi, 0)
                try:
                    value = eval(formula)
                except:
                    value = None
                table.append({'KPI': kpi, 'Before': B, 'After': A, 'Compare (%)': value})
            df_table = pd.DataFrame(table)
            st.dataframe(df_table, use_container_width=True)
            # download
            buf = BytesIO()
            df_table.to_excel(buf, index=False)
            buf.seek(0)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.download_button(
                label="Download Table as Excel",
                data=buf,
                file_name=f"KPI_Table_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # --- KPI Chart ---
    st.sidebar.header("3. KPI Chart")
    kpi_list = [k for k in raw_formulas.keys()]
    selected_kpi = st.sidebar.selectbox("Select KPI for Chart", kpi_list)
    if st.sidebar.button("Show Chart") and selected_kpi:
        df_chart = df.groupby('Day')[selected_kpi].sum().reset_index()
        df_chart['Date'] = pd.to_datetime(df_chart['Day'], format='%d/%m/%Y')
        df_chart = df_chart.sort_values('Date')
        fig, ax = plt.subplots(figsize=(8,4))
        line, = ax.plot(df_chart['Date'], df_chart[selected_kpi], marker='o')
        ax.set_title(f"{selected_kpi} over Time")
        ax.set_xlabel("Date")
        ax.set_ylabel(selected_kpi)
        plt.xticks(rotation=45)
        st.pyplot(fig)
        # download PNG
        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
        buf.seek(0)
        timest = datetime.now().strftime('%Y%m%d_%H%M%S')
        st.download_button(
            label="Download Chart PNG",
            data=buf,
            file_name=f"{selected_kpi.replace(' ','_')}_{timest}.png",
            mime="image/png"
        )
else:
    st.info("Upload input Performance Excel(s) and KPI_formula.xlsx to proceed.")

st.markdown("---")
st.caption("Powered by Streamlit")

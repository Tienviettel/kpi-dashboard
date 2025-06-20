import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import os

# Cấu hình trang
st.set_page_config(page_title="KPI Evaluation Dashboard", layout="wide")
st.title("KPI Evaluation Dashboard")

# Upload file dữ liệu và file công thức KPI
st.header("1. Upload Files")
input_files = st.file_uploader(
    "Performance Excel files (.xlsx)",
    type=["xlsx"], accept_multiple_files=True
)
formula_file = st.file_uploader(
    "KPI formula file (KPI sheet)",
    type=["xlsx"]
)

# Chỉ chạy khi đã upload đủ dữ liệu
if input_files and formula_file:
    # --- Đọc và gộp dữ liệu ---
    df_list = [pd.read_excel(f) for f in input_files]
    df_raw = pd.concat(df_list, ignore_index=True)
    # Tạo cột Day
    df_raw['Begin Time'] = pd.to_datetime(df_raw['Begin Time'], errors='coerce')
    df_raw['Day'] = df_raw['Begin Time'].dt.strftime('%d/%m/%Y')

    # --- Đọc file công thức KPI ---
    formulas_df = pd.read_excel(formula_file, sheet_name="KPI")
    # Danh sách KPI (cột từ thứ 2 trở đi)
    kpi_list = formulas_df.columns[1:].tolist()

    # --- Chọn ngày và KPI ---
    days = sorted(df_raw['Day'].dropna().unique(), key=lambda d: datetime.strptime(d,'%d/%m/%Y'))
    before = st.multiselect("Before Action Dates", days)
    after  = st.multiselect("After Action Dates", days)
    selected_kpi = st.selectbox("Select KPI to Chart", kpi_list)

    # Nút chạy phân tích và hiển thị song song
    if st.button("Run Analysis"):
        if not before or not after or not selected_kpi:
            st.warning("Vui lòng chọn đủ Before Dates, After Dates và KPI.")
        else:
            # --- Tính bảng so sánh ---
            table = []
            for kpi in kpi_list:
                # Lấy công thức từ DataFrame
                formula = formulas_df[kpi].dropna().iloc[0]

                # Tính tổng A, B
                B = df_raw[df_raw['Day'].isin(before)][kpi].sum() if kpi in df_raw.columns else 0
                A = df_raw[df_raw['Day'].isin(after)][kpi].sum()  if kpi in df_raw.columns else 0
                try:
                    result = eval(formula)
                except Exception:
                    result = None
                table.append({'KPI': kpi, 'Before': B, 'After': A, 'Compare (%)': result})
            df_table = pd.DataFrame(table)

            # --- Tính dữ liệu cho chart KPI đã chọn ---
            df_chart = (
                df_raw.groupby('Day')[selected_kpi]
                      .sum()
                      .reset_index()
            ) if selected_kpi in df_raw.columns else pd.DataFrame({'Day':[], selected_kpi:[]})
            df_chart['Date'] = pd.to_datetime(df_chart['Day'], format='%d/%m/%Y')
            df_chart = df_chart.sort_values('Date')

            # --- Hiển thị song song ---
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Comparison Table")
                st.dataframe(df_table, use_container_width=True)
                # Tải về Excel
                buf = BytesIO()
                df_table.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button(
                    "Download Table as Excel",
                    data=buf,
                    file_name=f"KPI_Table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col2:
                st.subheader(f"{selected_kpi} Chart")
                fig, ax = plt.subplots()
                ax.plot(df_chart['Date'], df_chart[selected_kpi], marker='o')
                ax.set_title(selected_kpi)
                ax.set_xlabel("Date")
                ax.set_ylabel(selected_kpi)
                plt.xticks(rotation=45)
                st.pyplot(fig)
                # Tải về PNG
                buf2 = BytesIO()
                fig.savefig(buf2, format='png', bbox_inches='tight', dpi=300)
                buf2.seek(0)
                st.download_button(
                    "Download Chart as PNG",
                    data=buf2,
                    file_name=f"{selected_kpi}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png"
                )
else:
    st.info("Upload Performance Management Excel(s) và KPI_formula.xlsx để bắt đầu.")

st.markdown("---")
st.caption("Powered by Streamlit")

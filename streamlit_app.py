import pandas as pd
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import glob
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from collections import OrderedDict
import os
from tkinter import filedialog, messagebox
import mplcursors

# === Dữ liệu và hàm xử lý ban đầu ===
replace_dict_tdd = {
    "PRB Number Used on Downlink Channel": "LTE DL Physical Resource Block_Used",
    "PRB Number Available on Downlink Channel": "LTE DL Physical Resource Block_Available",
    "PRB Number Used on Uplink Channel": "LTE UL Physical Resource Block_Used",
    "PRB Number Available on Uplink Channel": "LTE UL Physical Resource Block_Available",
    "LTE Upload User Throughput_denumerator": "LTE Upload User Throughput_denumerato",
    "UL padding denumerator": "UL padding denominator"
}
def process_excel(file_path, is_tdd=False):
    xl = pd.ExcelFile(file_path)
    sheet_names = [s for s in xl.sheet_names if s.strip().lower() != "KPI(Counter)"]
    if not sheet_names:
        return pd.DataFrame()  # Không còn sheet nào để đọc
    df_main = xl.parse(sheet_names[0])
    df_main.columns = [col if not isinstance(col, str) else
                       col.replace("_FDD", "").replace("_TDD", "").replace(" FDD", "").replace(" TDD", "")
                       for col in df_main.columns]
    if is_tdd:
        df_main.rename(columns=replace_dict_tdd, inplace=True)
    return df_main

try:
    input_files = glob.glob("input/Performance Management*.xlsx")
    if not input_files:
        print("Không tìm thấy file nào trong thư mục 'input'. Vui lòng kiểm tra lại.")
    dfs = []
    for file_path in input_files:
        is_tdd = "TDD" in file_path.upper()
        df_temp = process_excel(file_path, is_tdd = is_tdd)
        #
        dfs.append(df_temp)
    if dfs:
        df = pd.concat(dfs, ignore_index=True)
    else:
        df = pd.DataFrame()
except Exception as e:
    print(f"Lỗi khi đọc file input: {e}")
    df = pd.DataFrame()

try:
    book = pd.ExcelFile("KPI_fomula.xlsx")
    df_sheet1 = book.parse("KPI")
except FileNotFoundError:
    print("Lỗi: Không tìm thấy tệp KPI_fomula.xlsx. Vui lòng kiểm tra lại.")
    df_sheet1 = pd.DataFrame()
except Exception as e:
    print(f"Lỗi khi đọc KPI_fomula.xlsx (sheet KPI): {e}")
    df_sheet1 = pd.DataFrame()

if not df.empty:
    if not df_sheet1.empty:
        df_sheet1.columns = [str(col).strip() for col in df_sheet1.columns]

    if 'Begin Time' in df.columns:
        df['Begin Time'] = pd.to_datetime(df['Begin Time'], errors='coerce')
        df['Day'] = df['Begin Time'].dt.strftime('%d/%m/%Y')
        df['Time'] = df['Begin Time'].dt.strftime('%H:%M:%S')
    else:
        df['Begin Time'] = pd.NaT
        df['Day'] = None
        df['Time'] = None

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = (df[col].astype(str).str.replace(',', '', regex=False).str.strip()
                       .replace(['', 'nan', 'NaN', 'None', 'null', 'NA', 'NaT'], '0'))
        try:
            if col not in ['Begin Time', 'Day', 'Time']:
                if df[col].dtype != 'object' or '0' in df[col].unique():
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        except Exception:
            pass

# ĐẢM BẢO CÁC TÊN CỘT TRONG ĐÂY KHỚP 100% VỚI OUTPUT df.columns CỦA BẠN
# TẠO HÀM KPI THƯỜNG
#  1. Định nghĩa HÀM tính cho từng KPI
def calc_eps_cssr(d):
    try:
        if d.get("LTE RRC Setup Attempt", 0) == 0 or d.get("LTE S1 Signaling Setup Attempt", 0) == 0 or (d.get("LTE E-RAB Initial Setup Attempt", 0) + d.get("LTE E-RAB Additional Setup Attempt", 0)) == 0:
            return 0
        return (
            d.get("LTE RRC Setup Success", 0) / d.get("LTE RRC Setup Attempt", 1) *
            d.get("LTE S1 Signaling Setup Success", 0) / d.get("LTE S1 Signaling Setup Attempt", 1) *
            (d.get("LTE E-RAB Initial Setup Success", 0) + d.get("LTE E-RAB Additional Setup Success", 0)) /
            (d.get("LTE E-RAB Initial Setup Attempt", 1) + d.get("LTE E-RAB Additional Setup Attempt", 0))
        ) * 100
    except Exception:
        return 0

def calc_eps_cdr(d):
    try:
        if d.get("LTE Call Attempt", 0) == 0:
            return 0
        return d.get("LTE Call Drop", 0) / d.get("LTE Call Attempt", 1) * 100
    except Exception:
        return 0

def calc_tu_prb_dl(d):
    try:
        if d.get("LTE DL Physical Resource Block_Available", 0) == 0:
            return 0
        return d.get("LTE DL Physical Resource Block_Used", 0) / d.get("LTE DL Physical Resource Block_Available", 1) * 100
    except Exception:
        return 0

def calc_tu_prb_ul(d):
    try:
        if d.get("LTE UL Physical Resource Block_Available", 0) == 0:
            return 0
        return d.get("LTE UL Physical Resource Block_Used", 0) / d.get("LTE UL Physical Resource Block_Available", 1) * 100
    except Exception:
        return 0

def calc_csfb_cssr(d):
    try:
        if d.get("LTE CSFB Attempt", 0) == 0:
            return 0
        return d.get("LTE CSFB Success", 0) / d.get("LTE CSFB Attempt", 1) * 100
    except Exception:
        return 0

def calc_dl_pdcp_tcp_rtt(d):
    try:
        if d.get("Cell DL TCP Packets Average RTT(ms)_denominator", 0) == 0:
            return 0
        return d.get("Cell DL TCP Packets Average RTT(ms)_numerator", 0) / d.get("Cell DL TCP Packets Average RTT(ms)_denominator", 1)
    except Exception:
        return 0

def calc_dl_cell_throughput(d):
    try:
        if d.get("LTE Download Cell Throughput_denominator", 0) == 0:
            return 0
        return d.get("LTE Download Cell Throughput_numerator", 0) / d.get("LTE Download Cell Throughput_denominator", 1)
    except Exception:
        return 0

def calc_ul_cell_throughput(d):
    try:
        if d.get("LTE Upload Cell Throughput_denumerator", 0) == 0:
            return 0
        return d.get("LTE Upload Cell Throughput_numerator", 0) / d.get("LTE Upload Cell Throughput_denumerator", 1)
    except Exception:
        return 0

def calc_dl_user_throughput(d):
    try:
        if d.get("LTE Download User Throughput_denominator", 0) == 0:
            return 0
        return d.get("LTE Download User Throughput_numerator", 0) / d.get("LTE Download User Throughput_denominator", 1)
    except Exception:
        return 0

def calc_ul_user_throughput(d):
    try:
        if d.get("LTE Upload User Throughput_denumerato", 0) == 0:
            return 0
        return d.get("LTE Upload User Throughput_numerator", 0) / d.get("LTE Upload User Throughput_denumerato", 1)
    except Exception:
        return 0

def calc_dl_ca_throughput(d):
    try:
        if d.get("CA DL Throughput denominator", 0) == 0:
            return 0
        return d.get("CA DL Throughput numerator", 0) / d.get("CA DL Throughput denominator", 1)
    except Exception:
        return 0

def calc_intra_freq_hosr(d):
    try:
        if d.get("LTE Intra-Frequency Handover out Attempt", 0) == 0:
            return 0
        return d.get("LTE Intra-Frequency Handover out Success", 0) / d.get("LTE Intra-Frequency Handover out Attempt", 1) * 100
    except Exception:
        return 0

def calc_inter_freq_hosr(d):
    try:
        if d.get("LTE Inter-Frequency Handover out Attempt", 0) == 0:
            return 0
        return d.get("LTE Inter-Frequency Handover out Success", 0) / d.get("LTE Inter-Frequency Handover out Attempt", 1) * 100
    except Exception:
        return 0

def calc_cqi_average(d):
    try:
        if d.get("CQI_denumerator", 0) == 0:
            return 0
        return d.get("CQI_numerator", 0) / d.get("CQI_denumerator", 1)
    except Exception:
        return 0

def calc_mimo_tm4_utilisation(d):
    try:
        if d.get("%MIMO TB Utilisation TM4_denominator", 0) == 0:
            return 0
        return d.get("%MIMO TB Utilisation TM4_numerator", 0) / d.get("%MIMO TB Utilisation TM4_denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_user_thrput_4mbps(d):
    try:
        if d.get("%DL User Throughput greater than 4Mbps_denominator", 0) == 0:
            return 0
        return d.get("%DL User Throughput greater than 4Mbps_numerator", 0) / d.get("%DL User Throughput greater than 4Mbps_denominator", 1) * 100
    except Exception:
        return 0

def calc_ul_user_thrput_1mbps(d):
    try:
        if d.get("%UL User Throughput greater than 1Mbps_denominator", 0) == 0:
            return 0
        return d.get("%UL User Throughput greater than 1Mbps_numerator", 0) / d.get("%UL User Throughput greater than 1Mbps_denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_qpsk_modulation(d):
    try:
        if d.get("DL QPSK Modulation_Denominator", 0) == 0:
            return 0
        return d.get("DL QPSK Modulation_Nominator", 0) / d.get("DL QPSK Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_16qam_modulation(d):
    try:
        if d.get("DL 16QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("DL 16QAM Modulation_Nominator", 0) / d.get("DL 16QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_64qam_modulation(d):
    try:
        if d.get("DL 64QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("DL 64QAM Modulation_Nominator", 0) / d.get("DL 64QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_256qam_modulation(d):
    try:
        if d.get("DL 256QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("DL 256QAM Modulation_Nominator", 0) / d.get("DL 256QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_ul_qpsk_modulation(d):
    try:
        if d.get("UL QPSK Modulation_Denominator", 0) == 0:
            return 0
        return d.get("UL QPSK Modulation_Nominator", 0) / d.get("UL QPSK Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_ul_16qam_modulation(d):
    try:
        if d.get("UL 16QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("UL 16QAM Modulation_Nominator", 0) / d.get("UL 16QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_ul_64qam_modulation(d):
    try:
        if d.get("UL 64QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("UL 64QAM Modulation_Nominator", 0) / d.get("UL 64QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_ul_256qam_modulation(d):
    try:
        if d.get("UL 256QAM Modulation_Denominator", 0) == 0:
            return 0
        return d.get("UL 256QAM Modulation_Nominator", 0) / d.get("UL 256QAM Modulation_Denominator", 1) * 100
    except Exception:
        return 0

def calc_sinr_pusch_average(d):
    try:
        if d.get("SINR_PUSCH_Average_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUSCH_Average_numerator", 0) / d.get("SINR_PUSCH_Average_denominator", 1)
    except Exception:
        return 0

def calc_sinr_pusch_gt_minus2(d):
    try:
        if d.get("SINR_PUSCH>-2_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUSCH>-2_nominator", 0) / d.get("SINR_PUSCH>-2_denominator", 1) * 100
    except Exception:
        return 0

def calc_sinr_pusch_gt_6(d):
    try:
        if d.get("SINR_PUSCH>6_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUSCH>6_nominator", 0) / d.get("SINR_PUSCH>6_denominator", 1) * 100
    except Exception:
        return 0

def calc_sinr_pucch_average(d):
    try:
        if d.get("SINR_PUCCH_Average_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUCCH_Average_numerator", 0) / d.get("SINR_PUCCH_Average_denominator", 1)
    except Exception:
        return 0

def calc_sinr_pucch_gt_minus6(d):
    try:
        if d.get("SINR_PUCCH>-6_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUCCH>-6_nominator", 0) / d.get("SINR_PUCCH>-6_denominator", 1) * 100
    except Exception:
        return 0

def calc_sinr_pucch_gt_0(d):
    try:
        if d.get("SINR_PUCCH>0_denominator", 0) == 0:
            return 0
        return d.get("SINR_PUCCH>0_nominator", 0) / d.get("SINR_PUCCH>0_denominator", 1) * 100
    except Exception:
        return 0

def calc_rsrp_average(d):
    try:
        if d.get("[ITBBU]RSRP_average_denominator", 0) == 0:
            return 0
        return d.get("[ITBBU]RSRP_average_numerator", 0) / d.get("[ITBBU]RSRP_average_denominator", 1)
    except Exception:
        return 0

def calc_rsrp_gt_minus120(d):
    try:
        if d.get("[ITBBU]%RSRP>-120_denominator", 0) == 0:
            return 0
        return d.get("[ITBBU]%RSRP>-120_nominator", 0) / d.get("[ITBBU]%RSRP>-120_denominator", 1) * 100
    except Exception:
        return 0

def calc_rsrp_gt_minus116(d):
    try:
        if d.get("[ITBBU]%RSRP>-116_denominator", 0) == 0:
            return 0
        return d.get("[ITBBU]%RSRP>-116_nominator", 0) / d.get("[ITBBU]%RSRP>-116_denominator", 1) * 100
    except Exception:
        return 0

def calc_rsrp_gt_minus106(d):
    try:
        if d.get("[ITBBU]%RSRP>-106_denominator", 0) == 0:
            return 0
        return d.get("[ITBBU]%RSRP>-106_nominator", 0) / d.get("[ITBBU]%RSRP>-106_denominator", 1) * 100
    except Exception:
        return 0

def calc_error_dl_tb(d):
    try:
        if d.get("Total Number of DL TBs", 0) == 0:
            return 0
        return d.get("Error Number of DL TBs", 0) / d.get("Total Number of DL TBs", 1) * 100
    except Exception:
        return 0

def calc_error_ul_tb(d):
    try:
        if d.get("Total Number of UL TBs", 0) == 0:
            return 0
        return d.get("Error Number of UL TBs", 0) / d.get("Total Number of UL TBs", 1) * 100
    except Exception:
        return 0

def calc_ri_gte2(d):
    try:
        if d.get("RI>=2_Denominator", 0) == 0:
            return 0
        return d.get("RI>=2_numerator", 0) / d.get("RI>=2_Denominator", 1) * 100
    except Exception:
        return 0

def calc_dl_se(d):
    try:
        if d.get("DL spectrum efficiency denominator", 0) == 0:
            return 0
        return d.get("DL spectrum efficiency numerator", 0) / d.get("DL spectrum efficiency denominator", 1)
    except Exception:
        return 0

def calc_ul_se(d):
    try:
        if d.get("UL spectrum efficiency denominator", 0) == 0:
            return 0
        return d.get("UL spectrum efficiency numerator", 0) / d.get("UL spectrum efficiency denominator", 1)
    except Exception:
        return 0

def calc_rsrq_average(d):
    try:
        if d.get("RSRQ Denominator", 0) == 0:
            return 0
        return d.get("RSRQ Numerator", 0) / d.get("RSRQ Denominator", 1)
    except Exception:
        return 0

def calc_rsrq_gt_minus12(d):
    try:
        if d.get("%RSRQ>-12_Denominator", 0) == 0:
            return 0
        return d.get("%RSRQ>-12_Numerator", 0) / d.get("%RSRQ>-12_Denominator", 1) * 100
    except Exception:
        return 0

def calc_rsrq_gt_minus15(d):
    try:
        if d.get("%RSRQ>-15_Denominator", 0) == 0:
            return 0
        return d.get("%RSRQ>-15_Numerator", 0) / d.get("%RSRQ>-15_Denominator", 1) * 100
    except Exception:
        return 0

def calc_volte_cdr(d):
    try:
        if d.get("[VoLTE] Call Drop Denumerator", 0) == 0:
            return 0
        return d.get("[VoLTE] Call Drop Numerator", 0) / d.get("[VoLTE] Call Drop Denumerator", 1) * 100
    except Exception:
        return 0

def calc_srvcc_sr(d):
    try:
        if d.get("[VoLTE] SRVCC Attempt", 0) == 0:
            return 0
        return d.get("[VoLTE] SRVCC Success", 0) / d.get("[VoLTE] SRVCC Attempt", 1) * 100
    except Exception:
        return 0

# 2. Tạo dict map tên KPI -> hàm

kpi_formulas = {
    "ePS-CSSR": calc_eps_cssr,
    "ePS CDR": calc_eps_cdr,
    "TU PRB DL": calc_tu_prb_dl,
    "TU PRB UL": calc_tu_prb_ul,
    "CSFB CSSR": calc_csfb_cssr,
    "DL PDCP TCP RTT": calc_dl_pdcp_tcp_rtt,
    "DL Cell Throughput": calc_dl_cell_throughput,
    "UL Cell Throughput": calc_ul_cell_throughput,
    "DL User Throughput": calc_dl_user_throughput,
    "UL User Throughput": calc_ul_user_throughput,
    "DL CA Throughput": calc_dl_ca_throughput,
    "Intra-Freq HOSR": calc_intra_freq_hosr,
    "Inter-Freq HOSR": calc_inter_freq_hosr,
    "CQI Average": calc_cqi_average,
    "%MIMO TM4 Utilisation": calc_mimo_tm4_utilisation,
    "%DL User Thrput >4Mbps": calc_dl_user_thrput_4mbps,
    "%UL User Thrput >1Mbps": calc_ul_user_thrput_1mbps,
    "%QPSK DL": calc_dl_qpsk_modulation,
    "%16QAM DL": calc_dl_16qam_modulation,
    "%64QAM DL": calc_dl_64qam_modulation,
    "%256QAM DL": calc_dl_256qam_modulation,
    "%QPSK UL": calc_ul_qpsk_modulation,
    "%16QAM UL": calc_ul_16qam_modulation,
    "%64QAM UL": calc_ul_64qam_modulation,
    "%256QAM UL": calc_ul_256qam_modulation,
    "SINR PUSCH Average": calc_sinr_pusch_average,
    "%SINR_PUSCH>-2": calc_sinr_pusch_gt_minus2,
    "%SINR_PUSCH>6": calc_sinr_pusch_gt_6,
    "SINR PUCCH Average": calc_sinr_pucch_average,
    "%SINR_PUCCH>-6": calc_sinr_pucch_gt_minus6,
    "%SINR_PUCCH>0": calc_sinr_pucch_gt_0,
    "RSRP average": calc_rsrp_average,
    "%RSRP>-120": calc_rsrp_gt_minus120,
    "%RSRP>-116": calc_rsrp_gt_minus116,
    "%RSRP>-106": calc_rsrp_gt_minus106,
    "%Error DL TB": calc_error_dl_tb,
    "%Error UL TB": calc_error_ul_tb,
    "%RI>=2": calc_ri_gte2,
    "DL SE": calc_dl_se,
    "UL SE": calc_ul_se,
    "RSRQ average": calc_rsrq_average,
    "%RSRQ>-12": calc_rsrq_gt_minus12,
    "%RSRQ>-15": calc_rsrq_gt_minus15,
    "%VoLTE CDR": calc_volte_cdr,
    "%SRVCC SR": calc_srvcc_sr
}
kpi_columns = {
    "ePS-CSSR": [
        "LTE RRC Setup Success", "LTE RRC Setup Attempt",
        "LTE S1 Signaling Setup Success", "LTE S1 Signaling Setup Attempt",
        "LTE E-RAB Initial Setup Success", "LTE E-RAB Additional Setup Success",
        "LTE E-RAB Initial Setup Attempt", "LTE E-RAB Additional Setup Attempt"
    ],
    "ePS CDR": ["LTE Call Drop", "LTE Call Attempt"],
    "TU PRB DL": ["LTE DL Physical Resource Block_Used", "LTE DL Physical Resource Block_Available"],
    "TU PRB UL": ["LTE UL Physical Resource Block_Used", "LTE UL Physical Resource Block_Available"],
    "CSFB CSSR": ["LTE CSFB Success", "LTE CSFB Attempt"],
    "DL PDCP TCP RTT": ["Cell DL TCP Packets Average RTT(ms)_numerator", "Cell DL TCP Packets Average RTT(ms)_denominator"],
    "DL Cell Throughput": ["LTE Download Cell Throughput_numerator", "LTE Download Cell Throughput_denominator"],
    "UL Cell Throughput": ["LTE Upload Cell Throughput_numerator", "LTE Upload Cell Throughput_denumerator"],
    "DL User Throughput": ["LTE Download User Throughput_numerator", "LTE Download User Throughput_denominator"],
    "UL User Throughput": ["LTE Upload User Throughput_numerator", "LTE Upload User Throughput_denumerato"],
    "DL CA Throughput": ["CA DL Throughput numerator", "CA DL Throughput denominator"],
    "Intra-Freq HOSR": ["LTE Intra-Frequency Handover out Success", "LTE Intra-Frequency Handover out Attempt"],
    "Inter-Freq HOSR": ["LTE Inter-Frequency Handover out Success", "LTE Inter-Frequency Handover out Attempt"],
    "CQI Average": ["CQI_numerator", "CQI_denumerator"],
    "%MIMO TM4 Utilisation": ["%MIMO TB Utilisation TM4_numerator", "%MIMO TB Utilisation TM4_denominator"],
    "%DL User Thrput >4Mbps": ["%DL User Throughput greater than 4Mbps_numerator", "%DL User Throughput greater than 4Mbps_denominator"],
    "%UL User Thrput >1Mbps": ["%UL User Throughput greater than 1Mbps_numerator", "%UL User Throughput greater than 1Mbps_denominator"],
    "%QPSK DL": ["DL QPSK Modulation_Nominator", "DL QPSK Modulation_Denominator"],
    "%16QAM DL": ["DL 16QAM Modulation_Nominator", "DL 16QAM Modulation_Denominator"],
    "%64QAM DL": ["DL 64QAM Modulation_Nominator", "DL 64QAM Modulation_Denominator"],
    "%256QAM DL": ["DL 256QAM Modulation_Nominator", "DL 256QAM Modulation_Denominator"],
    "%QPSK UL": ["UL QPSK Modulation_Nominator", "UL QPSK Modulation_Denominator"],
    "%16QAM UL": ["UL 16QAM Modulation_Nominator", "UL 16QAM Modulation_Denominator"],
    "%64QAM UL": ["UL 64QAM Modulation_Nominator", "UL 64QAM Modulation_Denominator"],
    "%256QAM UL": ["UL 256QAM Modulation_Nominator", "UL 256QAM Modulation_Denominator"],
    "SINR PUSCH Average": ["SINR_PUSCH_Average_numerator", "SINR_PUSCH_Average_denominator"],
    "%SINR_PUSCH>-2": ["SINR_PUSCH>-2_nominator", "SINR_PUSCH>-2_denominator"],
    "%SINR_PUSCH>6": ["SINR_PUSCH>6_nominator", "SINR_PUSCH>6_denominator"],
    "SINR PUCCH Average": ["SINR_PUCCH_Average_numerator", "SINR_PUCCH_Average_denominator"],
    "%SINR_PUCCH>-6": ["SINR_PUCCH>-6_nominator", "SINR_PUCCH>-6_denominator"],
    "%SINR_PUCCH>0": ["SINR_PUCCH>0_nominator", "SINR_PUCCH>0_denominator"],
    "RSRP average": ["[ITBBU]RSRP_average_numerator", "[ITBBU]RSRP_average_denominator"],
    "%RSRP>-120": ["[ITBBU]%RSRP>-120_nominator", "[ITBBU]%RSRP>-120_denominator"],
    "%RSRP>-116": ["[ITBBU]%RSRP>-116_nominator", "[ITBBU]%RSRP>-116_denominator"],
    "%RSRP>-106": ["[ITBBU]%RSRP>-106_nominator", "[ITBBU]%RSRP>-106_denominator"],
    "%Error DL TB": ["Error Number of DL TBs", "Total Number of DL TBs"],
    "%Error UL TB": ["Error Number of UL TBs", "Total Number of UL TBs"],
    "%RI>=2": ["RI>=2_numerator", "RI>=2_Denominator"],
    "DL SE": ["DL spectrum efficiency numerator", "DL spectrum efficiency denominator"],
    "UL SE": ["UL spectrum efficiency numerator", "UL spectrum efficiency denominator"],
    "RSRQ average": ["RSRQ Numerator", "RSRQ Denominator"],
    "%RSRQ>-12": ["%RSRQ>-12_Numerator", "%RSRQ>-12_Denominator"],
    "%RSRQ>-15": ["%RSRQ>-15_Numerator", "%RSRQ>-15_Denominator"],
    "%VoLTE CDR": ["[VoLTE] Call Drop Numerator", "[VoLTE] Call Drop Denumerator"],
    "%SRVCC SR": ["[VoLTE] SRVCC Success", "[VoLTE] SRVCC Attempt"]
}

all_raw_columns_needed_for_formulas = set()
if not df.empty:
    for kpi_name in kpi_formulas.keys():
        if kpi_name in kpi_columns:
            all_raw_columns_needed_for_formulas.update(kpi_columns[kpi_name])
    all_raw_columns_needed_for_formulas = [col for col in all_raw_columns_needed_for_formulas if col in df.columns]
else:
    all_raw_columns_needed_for_formulas = []

# ĐOẠN TÍNH KPI CHỈ CẦN SỬA THÀNH HÀM THƯỜNG
if not df.empty:
    kpi_results = {}
    for new_col, formula_func in kpi_formulas.items():
        try:
            result_series = df.apply(lambda row: formula_func(row.fillna(0)), axis=1)
            kpi_results[new_col] = result_series.fillna(0)
        except Exception as e:
            kpi_results[new_col] = pd.NA

    if kpi_results:
        df_kpi = pd.DataFrame(kpi_results)
        df = pd.concat([df.reset_index(drop=True), df_kpi.reset_index(drop=True)], axis=1)

if not df.empty:
    all_raw_columns_needed_for_formulas = [col for col in all_raw_columns_needed_for_formulas if col in df.columns]
else:
    all_raw_columns_needed_for_formulas = []

kpi_list = []
kpi_config = {}

if not df_sheet1.empty and df_sheet1.shape[1] > 1:
    kpi_list_raw = [str(col).strip() for col in df_sheet1.columns[1:] if str(col).strip() != ""]
    kpi_list = list(OrderedDict.fromkeys(kpi_list_raw))

if kpi_list:
    for kpi in kpi_list:
        formula_str = ""
        if kpi in df_sheet1.columns and not df_sheet1[kpi].empty:
            formula_str_series = df_sheet1[kpi].dropna()
            if not formula_str_series.empty:
                formula_str = str(formula_str_series.iloc[0]).strip()
        kpi_config[kpi] = {'formula': formula_str if formula_str else '((A-B)/B)*100'}

L1_items_global = ['LTE PS Traffic (GB)', '[VoLTE] Traffic (ERL)', 'TU PRB DL', 'TU PRB UL', 'DL User Throughput',
                   'UL User Throughput', 'ePS-CSSR', 'ePS CDR', 'Maximum Number of RRC Connection User', 'CQI Average',
                   '%RI>=2',
                   'Intra-Freq HOSR', 'Inter-Freq HOSR', 'RSRP average', 'RSRQ average', 'DL Cell Throughput',
                   'UL Cell Throughput',
                   'CA Traffic (GB)']
L2_items_global = ['CSFB CSSR', '%DL User Thrput >4Mbps', '%UL User Thrput >1Mbps',
                   'Downlink Maximum Active 2CC User Number of Pcell',
                   'Downlink Maximum Active 3CC User Number of Pcell',
                   'Uplink Maximum User Number of Scell', 'Average DL Active E-RAB Number',
                   'Maximum Cell DL PDCP Throughput(Mbps)',
                   '%MIMO TM4 Utilisation', 'DL PDCP TCP RTT', '%Error DL TB',
                   '%Error UL TB', 'DL CA Throughput',
                   'SINR PUSCH Average', 'SINR PUCCH Average', 'DL SE', 'UL SE', '%SINR_PUCCH>0']

L2_first_9_global = L2_items_global[:9]
L2_last_9_global = L2_items_global[9:]
L3_items_global = ['%SINR_PUCCH>-6', '%SINR_PUSCH>-2', '%SINR_PUSCH>6', '%RSRP>-120',
                   '%RSRP>-116', '%RSRP>-106',
                   '%RSRQ>-12', '%RSRQ>-15', '%QPSK UL', '%16QAM UL', '%64QAM UL', '%QPSK DL', '%16QAM DL', '%64QAM DL',
                   '%256QAM DL','%VoLTE CDR', '%SRVCC SR']

display_to_original_map_global = {
    "LTE PS Traffic (GB)": "LTE PS Traffic",
    "CA Traffic (GB)": "CA Traffic(MB)",  # Tên gốc trong kpi_list là "CA Traffic(MB)"
}

if kpi_list:
    for kpi_name_original in kpi_list:
        if kpi_name_original not in display_to_original_map_global.values():
            combined_L_items = L1_items_global + L2_items_global + L3_items_global
            if kpi_name_original in combined_L_items:
                display_to_original_map_global[kpi_name_original] = kpi_name_original


# === GIAO DIỆN === (Copy từ phiên bản trước, đảm bảo sử dụng _global cho các list)
root = tk.Tk()
root.title("KPI Evaluation Tool_Tiendv24")
root.geometry("1600x900")
root.configure(bg='white')

style = ttk.Style()
style.theme_use("clam")
style.configure('.', background='white', foreground='black', font=('Arial', 9))
style.configure('TFrame', background='white')
style.configure('TLabel', background='white', foreground='black', padding=2)
style.configure('TButton', padding=5)
style.configure('TCombobox', padding=5, fieldbackground='white', selectbackground='lightblue', selectforeground='black')
style.configure('TLabelframe', background='white', bordercolor='gray', padding=10)
style.configure('TLabelframe.Label', background='white', foreground='black', font=('Arial', 10, 'bold'))
style.configure("Treeview", rowheight=20, borderwidth=1, relief="solid", fieldbackground="white", background="white")
style.configure("Treeview.Heading", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", background="lightgray",
                foreground="black")
style.map("Treeview", background=[('selected', '#cce5ff')])
style.configure('White.TFrame', background='white')

frame_select = ttk.Frame(root, padding="10")
frame_select.pack(side="left", fill="y", padx=10, pady=10)

sorted_days = []
if 'Day' in df.columns and not df['Day'].isnull().all():
    unique_days = df['Day'].dropna().unique()
    try:
        sorted_days = sorted(unique_days, key=lambda x: datetime.strptime(x, "%d/%m/%Y"))
    except ValueError:
        print(f"Cảnh báo: Một số ngày không đúng định dạng '%d/%m/%Y'. Sắp xếp theo chuỗi.")
        sorted_days = sorted(list(unique_days))

listbox_label_b = ttk.Label(frame_select, text="Before Action Dates:")
listbox_label_b.pack(pady=(0, 2))
listbox_b = tk.Listbox(frame_select, selectmode="multiple", exportselection=False, height=12, width=20,
                       font=('Arial', 9), bg='white', fg='black', highlightbackground='gray', highlightcolor='blue',
                       selectbackground='lightblue')
for day_val in sorted_days: listbox_b.insert(tk.END, day_val)
listbox_b.pack(pady=(0, 10))

listbox_label_a = ttk.Label(frame_select, text="After Action Dates:")
listbox_label_a.pack(pady=(0, 2))
listbox_a = tk.Listbox(frame_select, selectmode="multiple", exportselection=False, height=12, width=20,
                       font=('Arial', 9), bg='white', fg='black', highlightbackground='gray', highlightcolor='blue',
                       selectbackground='lightblue')
for day_val in sorted_days: listbox_a.insert(tk.END, day_val)
listbox_a.pack(pady=(0, 10))

frame_main_content_right = ttk.Frame(root, padding="10")
frame_main_content_right.pack(side="right", fill="both", expand=True)
frame_main_content_right.rowconfigure(0, weight=2)
frame_main_content_right.rowconfigure(1, weight=1)
frame_main_content_right.columnconfigure(0, weight=1)

frame_kpi_display_area = ttk.LabelFrame(frame_main_content_right, text="KPI Evaluation Day")
frame_kpi_display_area.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
frame_kpi_display_area.rowconfigure(0, weight=1)
frame_kpi_display_area.columnconfigure(0, weight=1)

frame_chart_display_area = ttk.LabelFrame(frame_main_content_right, text="Chart KPI")
frame_chart_display_area.grid(row=1, column=0, sticky="nsew", padx=5, pady=(10, 5))
frame_chart_display_area.rowconfigure(0, weight=1)
frame_chart_display_area.columnconfigure(0, weight=1)

chart_canvas_global = None


def match_kpi_name(kpi_name_to_match, df_columns_list):
    k_norm_search = str(kpi_name_to_match).lower().replace('>=', '>=').replace('>', '>').replace('=', '').replace('%',
                                                                                                                  '').replace(
        ' ', '').replace('_', '').replace('.', '').replace('(', '').replace(')', '')
    base_k_norm_search = k_norm_search
    is_gb_search = False
    if k_norm_search.endswith("(gb)"):
        base_k_norm_search = k_norm_search[:-4]
        is_gb_search = True

    for col_original_df in df_columns_list:
        c_norm_df = str(col_original_df).lower().replace('>=', '>=').replace('>', '>').replace('=', '').replace('%',
                                                                                                                '').replace(
            ' ', '').replace('_', '').replace('.', '').replace('(', '').replace(')', '')
        if is_gb_search:
            if c_norm_df.endswith("(mb)"):
                base_c_norm_df = c_norm_df[:-4]
                if base_k_norm_search == base_c_norm_df: return col_original_df
            elif base_k_norm_search == c_norm_df:
                return col_original_df
        else:
            if k_norm_search == c_norm_df: return col_original_df
    return None


def on_compare():
    global chart_canvas_global, all_raw_columns_needed_for_formulas
    global L1_items_global, L2_items_global, L3_items_global, L2_first_9_global, L2_last_9_global, display_to_original_map_global

    if chart_canvas_global:
        chart_canvas_global.get_tk_widget().destroy()
        chart_canvas_global = None

    selected_days_a = [listbox_a.get(i) for i in listbox_a.curselection()]
    selected_days_b = [listbox_b.get(i) for i in listbox_b.curselection()]
    if not selected_days_a or not selected_days_b:
        tk.messagebox.showwarning("Chưa chọn ngày", "Vui lòng chọn ngày cho cả Before Action và After Action.")
        return
    if df.empty:
        tk.messagebox.showerror("Lỗi dữ liệu", "Không có dữ liệu để xử lý. Vui lòng kiểm tra file input.")
        return

    data_a = df[df['Day'].isin(selected_days_a)]
    data_b = df[df['Day'].isin(selected_days_b)]
    df_cols = df.columns.tolist()

    summed_data_a = pd.Series(dtype='float64')
    summed_data_b = pd.Series(dtype='float64')

    if not data_a.empty and all_raw_columns_needed_for_formulas:
        cols_to_sum_a = [col for col in all_raw_columns_needed_for_formulas if col in data_a.columns]
        if cols_to_sum_a: summed_data_a = data_a[cols_to_sum_a].sum(min_count=1)
    if not data_b.empty and all_raw_columns_needed_for_formulas:
        cols_to_sum_b = [col for col in all_raw_columns_needed_for_formulas if col in data_b.columns]
        if cols_to_sum_b: summed_data_b = data_b[cols_to_sum_b].sum(min_count=1)

    all_display_kpis = list(
        OrderedDict.fromkeys(L1_items_global + L2_first_9_global + L2_last_9_global + L3_items_global))
    if 'Type' in all_display_kpis: all_display_kpis.remove('Type')

    row_type_display_new = ["Type"] + all_display_kpis
    row_b_values = ["Before Action"] + [pd.NA] * len(all_display_kpis)
    row_a_values = ["After Action"] + [pd.NA] * len(all_display_kpis)
    row_diff_values = ["Compare (%)"] + [pd.NA] * len(all_display_kpis)

    for kpi_original_name in kpi_list:
        current_display_name = kpi_original_name
        for disp_name, orig_name in display_to_original_map_global.items():
            if orig_name == kpi_original_name:
                current_display_name = disp_name
                break

        if current_display_name not in all_display_kpis: continue
        try:
            value_idx = all_display_kpis.index(current_display_name)
        except ValueError:
            continue

        actual_col_name = match_kpi_name(kpi_original_name, df_cols)
        config = kpi_config.get(kpi_original_name, {'formula': '((A-B)/B)*100'})
        formula_eval_str = config['formula']
        value_a, value_b = pd.NA, pd.NA
        if kpi_original_name == "Maximum Cell DL PDCP Throughput(Mbps)":
            if actual_col_name and actual_col_name in data_a.columns and not data_a[actual_col_name].dropna().empty:
                value_a = data_a[actual_col_name].mean()
            if actual_col_name and actual_col_name in data_b.columns and not data_b[actual_col_name].dropna().empty:
                value_b = data_b[actual_col_name].mean()
        elif kpi_original_name in kpi_formulas:
            formula_to_apply = kpi_formulas[kpi_original_name]
            try:
                if not summed_data_a.empty and not summed_data_a.isnull().all():
                    required_cols_for_lambda = kpi_columns.get(kpi_original_name, [])
                    if all(col in summed_data_a.index and pd.notna(summed_data_a.get(col)) for col in
                           required_cols_for_lambda):
                        value_a = formula_to_apply(summed_data_a.fillna(0))
            except ZeroDivisionError:
                value_a = float('inf')
            except KeyError:
                pass
            except Exception:
                pass
            try:
                if not summed_data_b.empty and not summed_data_b.isnull().all():
                    required_cols_for_lambda = kpi_columns.get(kpi_original_name, [])

                    if all(col in summed_data_b.index and pd.notna(summed_data_b.get(col)) for col in
                           required_cols_for_lambda):
                        value_b = formula_to_apply(summed_data_b.fillna(0))
            except ZeroDivisionError:
                value_b = float('inf')
            except KeyError:
                pass
            except Exception:
                pass
        elif actual_col_name:
            if actual_col_name in data_a.columns and not data_a[actual_col_name].dropna().empty:
                value_a = data_a[actual_col_name].sum()
            if actual_col_name in data_b.columns and not data_b[actual_col_name].dropna().empty:
                value_b = data_b[actual_col_name].sum()

        if kpi_original_name == "CA Traffic(MB)" or kpi_original_name == "LTE PS Traffic":
            if pd.notna(value_a):
                try:
                    value_a = float(value_a) / 1000.0
                except:
                    value_a = pd.NA
            if pd.notna(value_b):
                try:
                    value_b = float(value_b) / 1000.0
                except:
                    value_b = pd.NA

        compare = pd.NA
        try:
            val_a_float = float(value_a) if pd.notna(value_a) and value_a != "" else None
            val_b_float = float(value_b) if pd.notna(value_b) and value_b != "" else None
            if val_b_float is not None and val_b_float != 0 and val_a_float is not None:
                A, B = val_a_float, val_b_float;
                compare = eval(formula_eval_str)
            elif val_a_float is not None and val_b_float is not None and val_b_float == 0 and val_a_float != 0:
                compare = float('inf')
            elif val_a_float is not None and val_b_float is not None and val_b_float == 0 and val_a_float == 0:
                compare = 0.0
            elif val_a_float is not None and val_b_float is None:
                compare = float('inf') if val_a_float > 0 else (float('-inf') if val_a_float < 0 else 0.0)
        except ZeroDivisionError:
            compare = float('inf') if (val_a_float is not None and val_a_float > 0) else (
                float('-inf') if (val_a_float is not None and val_a_float < 0) else 0.0)
        except Exception:
            compare = pd.NA

        row_b_values[value_idx + 1] = round(value_b, 2) if pd.notna(value_b) and value_b != "" else ""
        row_a_values[value_idx + 1] = round(value_a, 2) if pd.notna(value_a) and value_a != "" else ""
        if pd.isna(compare):
            display_compare = ""
        elif compare == float('inf'):
            display_compare = "Inf"
        elif compare == float('-inf'):
            display_compare = "-Inf"
        else:
            display_compare = round(compare, 2)
        row_diff_values[value_idx + 1] = display_compare

    for widget in frame_kpi_display_area.winfo_children(): widget.destroy()
    canvas_outer = tk.Canvas(frame_kpi_display_area, bg='white', highlightthickness=0)
    scroll_y = ttk.Scrollbar(frame_kpi_display_area, orient="vertical", command=canvas_outer.yview)
    scroll_x = ttk.Scrollbar(frame_kpi_display_area, orient="horizontal", command=canvas_outer.xview)
    canvas_outer.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
    canvas_outer.grid(row=0, column=0, sticky="nsew")
    scroll_y.grid(row=0, column=1, sticky="ns")
    scroll_x.grid(row=1, column=0, sticky="ew")
    outer_frame_for_table = ttk.Frame(canvas_outer, style='White.TFrame')
    canvas_outer.create_window((0, 0), window=outer_frame_for_table, anchor="nw")
    outer_frame_for_table.bind("<Configure>",
                               lambda event: canvas_outer.configure(scrollregion=canvas_outer.bbox("all")))

    headers_sets_display = [
        ["Type"] + L1_items_global + L2_first_9_global,
        ["Type"] + L2_last_9_global + L3_items_global
    ]
    rows_to_display_final = [row_type_display_new, row_b_values, row_a_values, row_diff_values]

    for headers_subset_display_names in headers_sets_display:
        sub_table_frame = ttk.Frame(outer_frame_for_table, relief="solid", borderwidth=1, style='White.TFrame')
        sub_table_frame.pack(side="top", fill="x", pady=5, padx=5, expand=True)
        for row_idx, current_row_data_list_final in enumerate(rows_to_display_final):
            for col_idx, kpi_header_for_subset_display_name in enumerate(headers_subset_display_names):
                try:
                    data_list_idx = row_type_display_new.index(kpi_header_for_subset_display_name)
                    value_to_display = current_row_data_list_final[data_list_idx]
                except (ValueError, IndexError):
                    value_to_display = ""
                fg_color, font_weight = "black", "normal"
                is_header_row = (row_idx == 0)
                original_kpi_name_for_color = display_to_original_map_global.get(kpi_header_for_subset_display_name,
                                                                                 kpi_header_for_subset_display_name)
                is_compare_row = (row_idx == 3 and kpi_header_for_subset_display_name != "Type")

                if is_header_row: font_weight = "bold"
                if is_compare_row:
                    font_weight = "bold"
                    try:
                        str_val_compare = str(value_to_display).strip().lower()
                        val_float_compare = pd.NA
                        if str_val_compare == "inf":
                            val_float_compare = float('inf')
                        elif str_val_compare == "-inf":
                            val_float_compare = float('-inf')
                        elif not str_val_compare or str_val_compare in ["none", "nan", "nat"]:
                            val_float_compare = pd.NA
                        else:
                            val_float_compare = float(str_val_compare)

                        if pd.isna(val_float_compare):
                            fg_color = "black"
                        elif original_kpi_name_for_color in ["RSRQ average", "RSRP average", "ePS CDR"]:
                            if val_float_compare == float('inf'):
                                fg_color = "red"
                            elif val_float_compare == float('-inf'):
                                fg_color = "green"
                            elif val_float_compare < 0:
                                fg_color = "green"
                            elif val_float_compare > 0:
                                fg_color = "red"
                            else:
                                fg_color = "black"
                        else:
                            formula_str_for_color = kpi_config.get(original_kpi_name_for_color, {}).get("formula", "")
                            is_inverse_kpi = "100-A" in formula_str_for_color or "100 - A" in formula_str_for_color
                            if val_float_compare == float('inf'):
                                fg_color = "red" if not is_inverse_kpi else "green"
                            elif val_float_compare == float('-inf'):
                                fg_color = "green" if not is_inverse_kpi else "red"
                            elif is_inverse_kpi:
                                if val_float_compare < 0:
                                    fg_color = "green"
                                elif val_float_compare > 0:
                                    fg_color = "red"
                                else:
                                    fg_color = "black"
                            else:
                                if val_float_compare > 0:
                                    fg_color = "green"
                                elif val_float_compare < 0:
                                    fg_color = "red"
                                else:
                                    fg_color = "black"
                    except (ValueError, TypeError):
                        if str(value_to_display).strip().lower() == "inf":
                            fg_color = "red"
                        elif str(value_to_display).strip().lower() == "-inf":
                            fg_color = "green"
                        else:
                            fg_color = "black"
                label = ttk.Label(sub_table_frame, text=str(value_to_display), relief="solid", borderwidth=1,
                                  padding=(6, 3), font=("Arial", 9, font_weight), foreground=fg_color, anchor="center",
                                  width=22, wraplength=160)
                label.grid(row=row_idx, column=col_idx, sticky="nsew")
        for col_idx_header in range(len(headers_subset_display_names)):
            sub_table_frame.grid_columnconfigure(col_idx_header, weight=1, minsize=100)


button_frame = ttk.Frame(frame_select)
button_frame.pack(pady=20, fill="x")
ttk.Button(button_frame, text="Calculate & Show Table", command=on_compare, style="Accent.TButton").pack(fill="x",
                                                                                                         pady=5)


def export_to_excel():
    global all_raw_columns_needed_for_formulas, display_to_original_map_global
    if df.empty: tk.messagebox.showerror("Lỗi dữ liệu", "Không có dữ liệu để xuất."); return
    selected_days_a = [listbox_a.get(i) for i in listbox_a.curselection()]
    selected_days_b = [listbox_b.get(i) for i in listbox_b.curselection()]
    if not selected_days_a or not selected_days_b: tk.messagebox.showwarning("Chưa chọn ngày",
                                                                             "Vui lòng chọn ngày để xuất Excel."); return
    data_a_export = df[df['Day'].isin(selected_days_a)]
    data_b_export = df[df['Day'].isin(selected_days_b)]
    summed_data_a_export, summed_data_b_export = pd.Series(dtype='float64'), pd.Series(dtype='float64')
    if not data_a_export.empty and all_raw_columns_needed_for_formulas:
        cols_to_sum_a = [col for col in all_raw_columns_needed_for_formulas if col in data_a_export.columns]
        if cols_to_sum_a: summed_data_a_export = data_a_export[cols_to_sum_a].sum(min_count=1)
    if not data_b_export.empty and all_raw_columns_needed_for_formulas:
        cols_to_sum_b = [col for col in all_raw_columns_needed_for_formulas if col in data_b_export.columns]
        if cols_to_sum_b: summed_data_b_export = data_b_export[cols_to_sum_b].sum(min_count=1)

    kpi_display_names_for_excel = []
    for k_orig in kpi_list:
        disp_name = k_orig
        for d, o in display_to_original_map_global.items():  # Sử dụng map global
            if o == k_orig: disp_name = d; break
        kpi_display_names_for_excel.append(disp_name)

    export_data = {"KPI": kpi_display_names_for_excel}
    before_values, after_values, compare_values_export = [], [], []
    df_cols_export = df.columns.tolist()

    for kpi_original_name_export in kpi_list:
        actual_col_export = match_kpi_name(kpi_original_name_export, df_cols_export)
        config_export = kpi_config.get(kpi_original_name_export, {'formula': '((A-B)/B)*100'})
        formula_eval_export = config_export['formula']
        val_a_e, val_b_e = pd.NA, pd.NA

        if kpi_original_name_export == "Maximum Cell DL PDCP Throughput(Mbps)":
            if actual_col_export and actual_col_export in data_a_export.columns and not data_a_export[
                actual_col_export].dropna().empty: val_a_e = data_a_export[actual_col_export].mean()
            if actual_col_export and actual_col_export in data_b_export.columns and not data_b_export[
                actual_col_export].dropna().empty: val_b_e = data_b_export[actual_col_export].mean()
        elif kpi_original_name_export in kpi_formulas:
            formula_to_apply = kpi_formulas[kpi_original_name_export]  # Sửa tên biến
            try:
                if not summed_data_a_export.empty and not summed_data_a_export.isnull().all():
                    required_cols = kpi_columns.get(kpi_original_name_export, [])
                    if all(col in summed_data_a_export.index and pd.notna(summed_data_a_export.get(col)) for col in
                           required_cols):
                        val_a_e = formula_to_apply(summed_data_a_export.fillna(0))
            except:
                val_a_e = pd.NA  # Đơn giản hóa xử lý lỗi cho export
            try:
                if not summed_data_b_export.empty and not summed_data_b_export.isnull().all():
                    required_cols = kpi_columns.get(kpi_original_name_export, [])

                    if all(col in summed_data_b_export.index and pd.notna(summed_data_b_export.get(col)) for col in
                           required_cols): val_b_e = formula_to_apply(summed_data_b_export.fillna(0))
            except:
                val_b_e = pd.NA
        elif actual_col_export:
            if actual_col_export in data_a_export.columns and not data_a_export[
                actual_col_export].dropna().empty: val_a_e = data_a_export[actual_col_export].sum()
            if actual_col_export in data_b_export.columns and not data_b_export[
                actual_col_export].dropna().empty: val_b_e = data_b_export[actual_col_export].sum()

        if kpi_original_name_export == "CA Traffic(MB)" or kpi_original_name_export == "LTE PS Traffic":
            if pd.notna(val_a_e):
                try:
                    val_a_e = float(val_a_e) / 1000.0;
                except: val_a_e = pd.NA
            if pd.notna(val_b_e):
                try: val_b_e = float(val_b_e) / 1000.0;
                except: val_b_e = pd.NA

        comp_e = pd.NA
        try:
            val_ae_float = float(val_a_e) if pd.notna(val_a_e) and val_a_e != "" else None
            val_be_float = float(val_b_e) if pd.notna(val_b_e) and val_b_e != "" else None
            if val_be_float is not None and val_be_float != 0 and val_ae_float is not None:
                A = val_ae_float
                B = val_be_float
                comp_e = eval(formula_eval_export)
            elif val_ae_float is not None and val_be_float is not None and val_be_float == 0 and val_ae_float != 0:
                comp_e = float('inf')
            elif val_ae_float is not None and val_be_float is not None and val_be_float == 0 and val_ae_float == 0:
                comp_e = 0.0
            elif val_ae_float is not None and val_be_float is None:
                comp_e = float('inf') if val_ae_float > 0 else (float('-inf') if val_ae_float < 0 else 0.0)
        except:
            comp_e = pd.NA
        before_values.append(round(val_b_e, 2) if pd.notna(val_b_e) and val_b_e != "" else None)
        after_values.append(round(val_a_e, 2) if pd.notna(val_a_e) and val_a_e != "" else None)
        if pd.isna(comp_e):
            display_comp_e = None
        elif comp_e == float('inf'):
            display_comp_e = "Inf"
        elif comp_e == float('-inf'):
            display_comp_e = "-Inf"
        else:
            display_comp_e = round(comp_e, 2)
        compare_values_export.append(display_comp_e)

    export_data["Before Action"], export_data["After Action"], export_data[
        "Compare (%)"] = before_values, after_values, compare_values_export;
    df_export = pd.DataFrame(export_data)
    # --- Tạo tên file có timestamp để tránh trùng ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # ví dụ "20250619_113045"
    output_path = f"KPI_Result_Export_{timestamp}.xlsx"
    try:
        df_export.to_excel(output_path, index=False); tk.messagebox.showinfo("Xuất thành công",
                                                                             f"Đã xuất KPI ra file: {output_path}")
    except Exception as e:
        tk.messagebox.showerror("Lỗi xuất file", f"Lỗi khi xuất file: {e}")


ttk.Button(button_frame, text="Export Table to Excel", command=export_to_excel).pack(fill="x", pady=5)

chart_controls_frame = ttk.Frame(frame_select)
chart_controls_frame.pack(pady=10, fill="x", side="bottom")
ttk.Label(chart_controls_frame, text="Pick KPI for Chart:").pack(pady=(10, 2))
chart_kpi_var = tk.StringVar()

all_display_kpis_for_chart_combo = []
temp_L_items_for_chart = L1_items_global + L2_items_global + L3_items_global
for kpi_disp_cand in list(OrderedDict.fromkeys(temp_L_items_for_chart)):
    original_name_for_match = display_to_original_map_global.get(kpi_disp_cand, kpi_disp_cand)
    if match_kpi_name(original_name_for_match, df.columns.tolist() if not df.empty else []):
        all_display_kpis_for_chart_combo.append(kpi_disp_cand)

chart_kpi_combo = ttk.Combobox(chart_controls_frame, textvariable=chart_kpi_var,
                               values=all_display_kpis_for_chart_combo, width=28, state="readonly")
chart_kpi_combo.pack(pady=(0, 5))
if all_display_kpis_for_chart_combo: chart_kpi_combo.current(0)


def draw_kpi_chart_on_canvas():
    global chart_canvas_global, display_to_original_map_global
    selected_kpi_display_name = chart_kpi_var.get()
    if not selected_kpi_display_name:
        return
    if df.empty:
        tk.messagebox.showerror("Lỗi dữ liệu", "Không có dữ liệu để vẽ biểu đồ.")
        return

    original_kpi_name_for_chart = display_to_original_map_global.get(
        selected_kpi_display_name, selected_kpi_display_name)

    df_grouped_result = []

    try:
        df['Day_dt'] = pd.to_datetime(df['Day'], format="%d/%m/%Y", errors='coerce')
    except:
        tk.messagebox.showerror("Lỗi", "Không thể chuyển đổi ngày")
        return

    df_valid = df.dropna(subset=['Day_dt'])

    # === Xử lý ngoại lệ đặc biệt ===
    if original_kpi_name_for_chart == "Maximum Cell DL PDCP Throughput(Mbps)":
        actual_col = match_kpi_name(original_kpi_name_for_chart, df.columns)
        if not actual_col:
            tk.messagebox.showwarning("Không tìm thấy KPI", f"KPI {selected_kpi_display_name} không có trong dữ liệu.")
            return

        for day_val, group_df in df_valid.groupby("Day_dt"):
            value = group_df[actual_col].mean()
            df_grouped_result.append((day_val, value))
    elif original_kpi_name_for_chart in kpi_formulas:
        formula_func = kpi_formulas[original_kpi_name_for_chart]
        required_cols = kpi_columns.get(original_kpi_name_for_chart, [])

        for day_val, group_df in df_valid.groupby("Day_dt"):
            summed = group_df[required_cols].sum()
            try:
                kpi_value = formula_func(summed.fillna(0))
                df_grouped_result.append((day_val, kpi_value))
            except:
                df_grouped_result.append((day_val, None))
    else:
        actual_col = match_kpi_name(original_kpi_name_for_chart, df.columns)
        if not actual_col:
            tk.messagebox.showwarning("Không tìm thấy KPI", f"KPI {selected_kpi_display_name} không có trong dữ liệu.")
            return

        for day_val, group_df in df_valid.groupby("Day_dt"):
            value = group_df[actual_col].sum()
            if original_kpi_name_for_chart in ["CA Traffic(MB)", "LTE PS Traffic"]:
                value = value / 1000.0
            df_grouped_result.append((day_val, value))

    # === Vẽ biểu đồ ===
    df_plot_ready = pd.DataFrame(df_grouped_result, columns=["Day", "KPI_Value"])
    df_plot_ready = df_plot_ready.sort_values("Day")

    if df_plot_ready.empty:
        tk.messagebox.showinfo("Không có dữ liệu",
                               f"Không có dữ liệu hợp lệ cho KPI: {selected_kpi_display_name} để vẽ biểu đồ.")
        return

    for widget in frame_chart_display_area.winfo_children():
        widget.destroy()

    fig, ax = plt.subplots(figsize=(10, 4), facecolor='white')
    ax.set_facecolor('white')

    # === Vẽ line chart
    line, = ax.plot(df_plot_ready["Day"].dt.strftime("%d/%m"), df_plot_ready["KPI_Value"],
                    marker='o', color='dodgerblue', linestyle='-', linewidth=2, markersize=6)

    # === Hover tooltip với khả năng ẩn khi click ra ngoài
    cursor = mplcursors.cursor(line, hover=True)

    # Tạo biến toàn cục lưu annotation đang hiển thị
    current_annotation = {'ann': None}

    def on_hover(sel):
        day_text = df_plot_ready['Day'].dt.strftime('%d/%m').iloc[sel.index]
        value = df_plot_ready['KPI_Value'].iloc[sel.index]
        sel.annotation.set_text(f"Day: {day_text}\nValue: {value:.2f}")
        current_annotation['ann'] = sel.annotation

    def on_click(event):
        # Nếu click không trúng điểm nào -> ẩn tooltip
        if current_annotation['ann'] is not None:
            current_annotation['ann'].set_visible(False)
            chart_canvas_global.draw()
            current_annotation['ann'] = None

    cursor.connect("add", on_hover)
    fig.canvas.mpl_connect("button_press_event", on_click)

    ax.set_title(f"Chart KPI: {selected_kpi_display_name}", fontsize=12, fontweight='bold', color='black')
    ax.set_xlabel("Day", fontsize=10, color='black')
    ax.set_ylabel(selected_kpi_display_name, fontsize=10, color='black')
    ax.tick_params(axis='x', rotation=45, labelsize=8, colors='black')
    ax.tick_params(axis='y', labelsize=8, colors='black')
    ax.grid(True, linestyle='--', alpha=0.7, color='lightgray')
    for spine in ax.spines.values():
        spine.set_edgecolor('black')
    plt.tight_layout()
    # Giữ figure và tên KPI để hàm save truy cập
    global current_fig, current_kpi_name
    current_fig = fig
    current_kpi_name = selected_kpi_display_name

    chart_canvas_global = FigureCanvasTkAgg(fig, master=frame_chart_display_area)
    chart_canvas_global.draw()
    chart_canvas_global.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    chart_canvas_global.get_tk_widget().configure(bg='white')

def save_current_chart():
    if not globals().get('current_fig'):
        messagebox.showwarning("Chưa có chart", "Bạn chưa hiển thị chart nào để lưu.")
        return

    default_name = current_kpi_name.replace(' ', '_') + '.png'
    file_path = filedialog.asksaveasfilename(
        defaultextension=".png",
        initialfile=default_name,
        filetypes=[("PNG files","*.png"),("All files","*.*")],
        title="Save chart as..."
    )
    if not file_path:
        return

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    current_fig.savefig(file_path, dpi=300, bbox_inches='tight')
    messagebox.showinfo("Đã lưu", f"Chart đã được lưu tại:\n{file_path}")

ttk.Button(chart_controls_frame, text="Display Chart", command=draw_kpi_chart_on_canvas, style="Accent.TButton").pack(
    fill="x", pady=5)
# Thêm nút Save Chart ngay bên dưới:
ttk.Button(
    chart_controls_frame,
    text="Save Chart",
    command=save_current_chart,
    style="Accent.TButton"
).pack(fill="x", pady=5)
style.configure("Accent.TButton", foreground="white", background="dodgerblue", font=('Arial', 9, 'bold'))
style.map("Accent.TButton", background=[('active', 'royalblue'), ('pressed', 'darkblue')],
          relief=[('pressed', 'sunken'), ('!pressed', 'raised')])

root.mainloop()

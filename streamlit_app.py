import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime

# --- Định nghĩa hàm tính KPI ---
def calc_eps_cssr(d):
    try:
        if d.get("LTE RRC Setup Attempt", 0) == 0 or d.get("LTE S1 Signaling Setup Attempt", 0) == 0 or (
            d.get("LTE E-RAB Initial Setup Attempt", 0) + d.get("LTE E-RAB Additional Setup Attempt", 0)
        ) == 0:
            return 0
        return (
            d.get("LTE RRC Setup Success", 0) / d.get("LTE RRC Setup Attempt", 1) *
            d.get("LTE S1 Signaling Setup Success", 0) / d.get("LTE S1 Signaling Setup Attempt", 1) *
            (d.get("LTE E-RAB Initial Setup Success", 0) + d.get("LTE E-RAB Additional Setup Success", 0)) /
            (d.get("LTE E-RAB Initial Setup Attempt", 1) + d.get("LTE E-RAB Additional Setup Attempt", 1))
        ) * 100
    except:
        return 0

def calc_eps_cdr(d):
    try:
        if d.get("LTE Call Attempt", 0) == 0:
            return 0
        return d.get("LTE Call Drop", 0) / d.get("LTE Call Attempt", 1) * 100
    except:
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

# --- Mapping tên KPI đến hàm tính ---
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

# --- Mapping KPI đến các cột raw cần tính ---
kpi_columns = {
    "ePS-CSSR": ["LTE RRC Setup Success", "LTE RRC Setup Attempt",
                  "LTE S1 Signaling Setup Success", "LTE S1 Signaling Setup Attempt",
                  "LTE E-RAB Initial Setup Success", "LTE E-RAB Additional Setup Success",
                  "LTE E-RAB Initial Setup Attempt", "LTE E-RAB Additional Setup Attempt"],
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

# --- Hàm load và clean dữ liệu ---
# --- Mapping cho TDD/FDD rename ---
replace_dict_tdd = {
    "PRB Number Used on Downlink Channel": "LTE DL Physical Resource Block_Used",
    "PRB Number Available on Downlink Channel": "LTE DL Physical Resource Block_Available",
    "PRB Number Used on Uplink Channel": "LTE UL Physical Resource Block_Used",
    "PRB Number Available on Uplink Channel": "LTE UL Physical Resource Block_Available",
    "LTE Upload User Throughput_denumerator": "LTE Upload User Throughput_denumerato",
    "UL padding denumerator": "UL padding denominator"
}

def process_excel_file(file_buffer, is_tdd=False):
    # Load Excel into DataFrame
    xl = pd.ExcelFile(file_buffer, engine='openpyxl')
    sheets = [s for s in xl.sheet_names if s.strip().lower() != "kpi(counter)"]
    if not sheets:
        return pd.DataFrame()
    df_main = xl.parse(sheets[0])
    # Clean column names
    df_main.columns = [col if not isinstance(col, str) else col.replace("_FDD", "").replace("_TDD", "").strip()
                       for col in df_main.columns]
    if is_tdd:
        df_main.rename(columns=replace_dict_tdd, inplace=True)
    return df_main

@st.cache_data
def load_data(files):
    dfs = []
    for f in files:
        # Determine TDD by filename
        is_tdd = 'TDD' in getattr(f, 'name', '').upper()
        # Use buffer or path
        df = process_excel_file(f, is_tdd=is_tdd)
        # Parse and clean
        if 'Begin Time' in df.columns:
            df['Begin Time'] = pd.to_datetime(df['Begin Time'], errors='coerce')
            df['Day'] = df['Begin Time'].dt.strftime('%d/%m/%Y')
        # Clean numeric columns
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = (df[col].astype(str)
                           .str.replace(',', '')
                           .replace(['', 'nan', 'None'], '0'))
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        dfs.append(df)
    # Concatenate all
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

# --- Giao diện Streamlit ---
st.set_page_config(page_title="KPI Dashboard", layout="wide")
st.title("KPI Evaluation Dashboard")

# Sidebar upload
with st.sidebar:
    st.header("Upload Files")
    perf_files   = st.file_uploader("Performance Excel (.xlsx)", type="xlsx", accept_multiple_files=True)
    formula_file = st.file_uploader("KPI formulas (.xlsx, sheet 'KPI')", type="xlsx")
perf_files = st.sidebar.file_uploader(
    "Excel files (.xlsx)", type="xlsx", accept_multiple_files=True
)

if perf_files and formula_file:
    # Load data và công thức
    df_raw = load_data(perf_files)
    formulas_df = pd.read_excel(formula_file, sheet_name='KPI', engine='openpyxl')
    # Danh sách KPI từ sheet
    kpi_list = [col for col in formulas_df.columns if col != formulas_df.columns[0]]

    # Chuẩn bị ngày
    days = (pd.to_datetime(df_raw['Day'], format='%d/%m/%Y', errors='coerce')
               .dropna().sort_values().dt.strftime('%d/%m/%Y').unique())

    # Sidebar chọn ngày và KPI
    with st.sidebar:
        st.header("Select Dates & KPI")
        before = st.multiselect("Before Dates", options=days)
        after  = st.multiselect("After Dates", options=days)
        selected_kpi = st.selectbox("Select KPI", options=kpi_list)
        run = st.button("Run Analysis")

    if run:
        # Tính bảng so sánh
        table_rows = []
        for kpi in kpi_list:
            func = kpi_formulas.get(kpi)
            cols = kpi_columns.get(kpi, [])
            # Tổng raw
            sum_b = df_raw[df_raw['Day'].isin(before)][cols].sum() if before else pd.Series(0)
            sum_a = df_raw[df_raw['Day'].isin(after)][cols].sum()  if after  else pd.Series(0)
            # Giá trị KPI
            val_b = func(sum_b) if func else 0
            val_a = func(sum_a) if func else 0
            # Công thức so sánh từ sheet
            formula_str = formulas_df[kpi].dropna().iloc[0]
            try:
                comp = eval(formula_str, {}, {'A': val_a, 'B': val_b})
            except Exception:
                comp = None
            table_rows.append({'KPI': kpi, 'Before': round(val_b,2), 'After': round(val_a,2), 'Compare (%)': round(comp,2) if comp is not None else None})
        df_table = pd.DataFrame(table_rows)

        # Tính dữ liệu chart
        chart_data = []
        func = kpi_formulas.get(selected_kpi)
        cols = kpi_columns.get(selected_kpi, [])
        for day in days:
            grp = df_raw[df_raw['Day'] == day]
            sums = grp[cols].sum()
            val  = func(sums) if func else 0
            chart_data.append({'Day': datetime.strptime(day, '%d/%m/%Y'), 'Value': val})
        df_chart = pd.DataFrame(chart_data).sort_values('Day')

        # Hiển thị song song
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Comparison Table")
            st.dataframe(df_table, use_container_width=True)
            buf = BytesIO(); df_table.to_excel(buf, index=False); buf.seek(0)
            st.download_button("Download Table Excel", buf,
                               file_name=f"KPI_Table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col2:
            st.subheader(f"{selected_kpi} Over Time")
            fig, ax = plt.subplots()
            ax.plot(df_chart['Day'], df_chart['Value'], marker='o')
            ax.set_xlabel('Date'); ax.set_ylabel(selected_kpi); plt.xticks(rotation=45)
            st.pyplot(fig)
            buf2 = BytesIO(); fig.savefig(buf2, format='png', bbox_inches='tight', dpi=300); buf2.seek(0)
            st.download_button("Download Chart PNG", buf2,
                               file_name=f"{selected_kpi}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                               mime="image/png")
else:
    st.info("Please upload Performance Excel(s) and KPI_formula.xlsx to begin.")


"""
PRS PriDE - Dashboard Overview (Streamlit version)
Jalankan dengan: streamlit run streamlit_app.py

Struktur:
- Sidebar: upload file Excel (opsional) untuk perbandingan data REAL vs dummy.
- 4 halaman: Overview, PRS Performance, AM Performance, Action Plan AM.
"""

import os
import base64
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="PRS PriDE — Dashboard Overview",
    page_icon="📊",
    layout="wide",
)

# ----------------------------------------------------------------------------
# THEME / CSS — mendukung mode Gelap & Terang, teal + glow tetap jadi ciri khasnya
# ----------------------------------------------------------------------------
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

THEMES = {
    "dark": {
        "CYAN": "#3fd6f0", "GOLD": "#f5b942", "GREEN": "#4fd18b", "RED": "#f0685f",
        "BG": "#06222e", "BG2": "#0a3244", "CARD": "#0f3f52", "MUTED": "#9fc4cf",
        "TEXT": "#f5fbfd", "HEADER_TEXT": "#04222c", "SIDEBAR_BG": "#08283a",
        "BORDER": "rgba(140,220,235,0.18)", "ROWBORDER": "rgba(255,255,255,0.06)",
        "TOTALBG": "rgba(255,255,255,0.09)", "GLOW1": "rgba(63,214,240,0.10)", "GLOW2": "rgba(245,185,66,0.08)",
        "GREEN_GLOW": "rgba(79,209,139,0.55)", "RED_GLOW": "rgba(240,104,95,0.55)",
    },
    "light": {
        "CYAN": "#0891b2", "GOLD": "#c2790a", "GREEN": "#0f9d58", "RED": "#d93025",
        "BG": "#eef7fa", "BG2": "#dcf0f5", "CARD": "#ffffff", "MUTED": "#5b7c8a",
        "TEXT": "#0a2f3f", "HEADER_TEXT": "#ffffff", "SIDEBAR_BG": "#ffffff",
        "BORDER": "rgba(8,60,75,0.15)", "ROWBORDER": "rgba(8,60,75,0.08)",
        "TOTALBG": "rgba(8,60,75,0.06)", "GLOW1": "rgba(8,145,178,0.07)", "GLOW2": "rgba(194,121,10,0.06)",
        "GREEN_GLOW": "rgba(15,157,88,0.45)", "RED_GLOW": "rgba(217,48,37,0.45)",
    },
}
T = THEMES[st.session_state.theme]
CYAN, GOLD, GREEN, RED = T["CYAN"], T["GOLD"], T["GREEN"], T["RED"]
BG, BG2, CARD, MUTED = T["BG"], T["BG2"], T["CARD"], T["MUTED"]
TEXT, HEADER_TEXT, SIDEBAR_BG = T["TEXT"], T["HEADER_TEXT"], T["SIDEBAR_BG"]
BORDER, ROWBORDER, TOTALBG = T["BORDER"], T["ROWBORDER"], T["TOTALBG"]
GLOW1, GLOW2 = T["GLOW1"], T["GLOW2"]
GREEN_GLOW, RED_GLOW = T["GREEN_GLOW"], T["RED_GLOW"]

st.markdown(f"""
<style>
.stApp {{
    background:
      radial-gradient(1200px 600px at 15% -10%, {GLOW1}, transparent 60%),
      radial-gradient(900px 500px at 100% 0%, {GLOW2}, transparent 55%),
      linear-gradient(180deg, {BG}, {BG2} 40%, {BG});
    color: {TEXT};
}}
[data-testid="stSidebar"] {{ background:{SIDEBAR_BG}; border-right:1px solid {BORDER}; }}
h1,h2,h3 {{ font-family:'Trebuchet MS', sans-serif; font-weight:800; color:{TEXT}; }}

.pill-title{{
    display:inline-block; padding:14px 46px; border-radius:999px;
    border:1.5px solid {CYAN}; background:{GLOW1};
    box-shadow:0 0 24px {GLOW1}, inset 0 0 18px {GLOW1};
    font-size:32px; font-weight:900; letter-spacing:2px; margin-bottom:6px; color:{TEXT};
}}
.subtitle{{ font-style:italic; color:{MUTED}; font-size:16px; margin-bottom:10px;}}

.card-box{{
    background:linear-gradient(160deg, {CARD}, {BG2});
    border:1px solid {BORDER}; border-radius:20px; padding:26px 16px;
    text-align:center; height:100%;
}}
.card-box .icon{{ font-size:42px; margin-bottom:10px; }}
.card-box h3{{ font-size:15px; letter-spacing:1px; margin-bottom:6px; color:{TEXT};}}
.card-box p{{ font-size:12px; color:{MUTED}; }}

.metric-card{{
    background:{CARD}; border:1px solid {BORDER}; border-radius:14px;
    padding:16px 18px; margin-bottom:6px;
}}
.metric-card .label{{ font-size:11px; color:{MUTED}; text-transform:uppercase; letter-spacing:1px; font-weight:600;}}
.metric-card .value{{ font-size:28px; font-weight:800; margin-top:4px; color:{TEXT};}}
.metric-card .delta.up{{ color:{GREEN}; font-size:12px; font-weight:600;}}
.metric-card .delta.down{{ color:{RED}; font-size:12px; font-weight:600;}}

div[data-testid="stMetric"]{{
    background:{CARD}; border:1px solid {BORDER}; border-radius:14px; padding:12px 16px;
}}
.stButton>button{{
    background:{CARD}; color:{TEXT}; border:1px solid {BORDER}; border-radius:10px;
    font-weight:700;
}}
.stButton>button:hover{{ border-color:{CYAN}; color:{CYAN}; }}

/* AM Performance scorecard style */
.am-banner{{
    background:{CYAN}; color:{HEADER_TEXT}; font-weight:900;
    font-size:24px; letter-spacing:1px; padding:20px 24px; border-radius:14px; margin:14px 0 18px;
    box-shadow:0 0 24px {GLOW1}; text-align:center;
}}
.am-panel{{ border:1px solid {BORDER}; border-radius:10px; overflow:hidden; margin-bottom:16px; background:{CARD}; }}
.am-header{{ background:{CYAN}; color:{HEADER_TEXT}; font-weight:800; padding:9px 14px; text-align:center; font-size:12.5px; letter-spacing:0.4px;}}
.am-header.gold{{ background:{GOLD}; }}
.am-subheader{{ background:{GLOW1}; color:{CYAN}; font-weight:700; padding:6px 14px; text-align:center; font-size:11.5px;}}
.am-subheader.gold{{ background:{GLOW2}; color:{GOLD}; }}
.am-row{{ display:flex; justify-content:space-between; padding:7px 14px; font-size:12.5px; border-bottom:1px solid {ROWBORDER};}}
.am-row span:first-child{{ color:{MUTED}; }}
.am-row span:last-child{{ font-weight:700; color:{TEXT}; }}
.am-row.g{{ background:rgba(15,157,88,0.12); }}
.am-row.g span:last-child{{ color:{GREEN}; }}
.am-row.r{{ background:rgba(217,48,37,0.12); }}
.am-row.r span:last-child{{ color:{RED}; }}
.am-row.total{{ background:{TOTALBG}; font-weight:800; }}
.am-row.total span{{ color:{TEXT}; }}

/* Badge ACH & GAP — lebih mencolok dari baris biasa */
.am-badge{{
    display:inline-block; padding:3px 12px; border-radius:999px;
    font-weight:900; font-size:12.5px; letter-spacing:0.3px;
}}
.am-badge.g{{ background:{GREEN}; color:#ffffff; box-shadow:0 0 12px {GREEN_GLOW}; }}
.am-badge.r{{ background:{RED}; color:#ffffff; box-shadow:0 0 12px {RED_GLOW}; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=MUTED, size=12),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    margin=dict(l=10, r=10, t=30, b=10),
)

# ----------------------------------------------------------------------------
# DUMMY DATA (fallback kalau belum upload data real)
# ----------------------------------------------------------------------------
def dummy_data():
    prs_monthly = pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun"],
        "Realisasi": [78, 82, 85, 88, 90, 92],
        "Target": [80, 80, 85, 85, 90, 90],
    })
    prs_kpi = pd.DataFrame({
        "Metric": ["Occupancy Rate", "Revenue Achievement", "Maintenance Completion", "Asset Utilization"],
        "Value": [87, 92, 78, 85],
        "Target": [85, 90, 85, 85],
    })
    prs_portfolio = pd.DataFrame({
        "Type": ["Office", "Residential", "Retail", "Industrial"],
        "Value": [38, 27, 20, 15],
    })
    prs_regional = pd.DataFrame({
        "Region": ["Jakarta", "Surabaya", "Bandung", "Medan", "Makassar"],
        "Value": [91, 84, 79, 73, 68],
    })
    am = pd.DataFrame({
        "Name": ["Rina Wulandari", "Bagus Santoso", "Dewi Anggraini", "Fajar Nugroho",
                 "Siti Rahmawati", "Andi Kurniawan", "Putri Handayani", "Yusuf Hidayat"],
        "Region": ["Jakarta", "Surabaya", "Bandung", "Medan", "Makassar", "Semarang", "Denpasar", "Palembang"],
        "Target": [96, 91, 88, 82, 79, 75, 73, 69],
        "Score": [94.2, 89.5, 86.1, 80.4, 77.8, 74.0, 71.6, 68.2],
        "Komunikasi": [92, 85, 88, 80, 82, 78, 75, 70],
        "Negosiasi": [88, 90, 80, 78, 81, 76, 74, 69],
        "Kepatuhan": [95, 80, 85, 79, 83, 77, 72, 68],
        "KepuasanKlien": [90, 88, 84, 81, 80, 75, 73, 70],
        "Pelaporan": [85, 82, 90, 77, 79, 74, 76, 66],
    })
    action = pd.DataFrame({
        "AM": ["Rina Wulandari", "Bagus Santoso", "Dewi Anggraini", "Fajar Nugroho",
               "Siti Rahmawati", "Andi Kurniawan", "Putri Handayani", "Yusuf Hidayat"],
        "Item": [
            "Follow-up renewal kontrak Gedung A", "Survey kepuasan tenant Q3",
            "Update laporan maintenance mingguan", "Negosiasi ulang kontrak vendor cleaning",
            "Audit aset idle regional Makassar", "Penyelesaian komplain tenant lantai 4",
            "Pembaruan data okupansi bulanan", "Koordinasi perbaikan lift Gedung C",
        ],
        "Due": ["2026-08-12", "2026-08-15", "2026-08-08", "2026-08-20",
                "2026-08-18", "2026-08-10", "2026-08-14", "2026-08-09"],
        "Priority": ["Tinggi", "Sedang", "Rendah", "Tinggi", "Sedang", "Tinggi", "Rendah", "Tinggi"],
        "Status": ["Berjalan", "Belum Mulai", "Selesai", "Berjalan", "Berjalan", "Belum Mulai", "Selesai", "Berjalan"],
    })
    return {"prs_monthly": prs_monthly, "prs_kpi": prs_kpi, "prs_portfolio": prs_portfolio,
            "prs_regional": prs_regional, "am": am, "action": action}


# ----------------------------------------------------------------------------
# AM PERFORMANCE — detail scorecard per Account Manager (dummy, 3 contoh).
# Untuk tambah AM lain, copy salah satu blok ini dan ganti angkanya.
# ----------------------------------------------------------------------------
AM_DETAIL_DEFAULT = {
    "Mochamad Faroq": {
        "rev_h1": {"real": "11,24M", "ach": 111},
        "lop_visit_h1": {"kecukupan_all_lop": 97, "jml_visit": 88, "target_visit": 96, "ach_visit": 92},
        "scal_h1": {"scal_bc": "3,38M", "net_scaling": "2,27M"},
        "rincian_h1": {"AO": "3,08M", "MO+": "0,10M", "TERMIN": "0,00M", "RO": "0,20M",
                       "SO": "-0,03M", "DO": "-0,95M", "MO-": "-0,03M", "ADJ": "-0,10M", "TOTAL": "2,27M"},
        "pacer_h1": {"result": 75, "process": 24, "total": 99, "kuadran": "KUADRAN 2"},
        "kecukupan_lop_juli": {"target_scal_rkap": "1,87M", "kebutuhan_lop": "3,74M", "est_rev_lop": "5,21M", "ach": 139},
        "visit_juli": {"target_cm": 16, "jml_visit_cm": 24, "ach_cm": 150,
                       "target_ytd": 112, "jml_visit_ytd": 112, "ach_ytd": 100},
        "scal_juli": {"scal_bc": "0,04M", "net_scaling": "-0,04M"},
        "rincian_juli": {"AO": "0,02M", "MO+": "0,02M", "TERMIN": "0,00M", "RO": "0,00M",
                         "SO": "0,00M", "DO": "-0,02M", "MO-": "0,00M", "ADJ": "-0,06M", "TOTAL": "-0,04M"},
        "kecukupan_agustus": {"est_rev_f3f4": "0,30M", "ach_f3f4": 8, "gap_f3f4": "3,44M",
                               "est_rev_all": "5,03M", "ach_all": 134, "gap_all": "-1,28M"},
        "visit_agustus": {"visit_cm": 0, "target_cm": 16, "ach_cm": 0,
                           "visit_ytd": 112, "target_ytd": 128, "ach_ytd": 88},
        "list_cc": [
            {"cc": "Adaro Indonesia PT", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Asian Agri Abadi Group", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Nusa Halmahera Mineral PT", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Riau Andalan Pulp & Paper", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Royal Golden Eagle Indonesia", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "PT Kalimantan Industrial Park Indonesia", "lop": "ADA LOP", "scal": "TANPA SCALING"},
            {"cc": "Saptaindra Sejati", "lop": "ADA LOP", "scal": "TANPA SCALING"},
        ],
    },
    "Rina Wulandari": {
        "rev_h1": {"real": "14,80M", "ach": 128},
        "lop_visit_h1": {"kecukupan_all_lop": 105, "jml_visit": 98, "target_visit": 96, "ach_visit": 102},
        "scal_h1": {"scal_bc": "4,10M", "net_scaling": "3,05M"},
        "rincian_h1": {"AO": "3,40M", "MO+": "0,25M", "TERMIN": "0,10M", "RO": "0,15M",
                       "SO": "-0,02M", "DO": "-0,60M", "MO-": "-0,01M", "ADJ": "-0,22M", "TOTAL": "3,05M"},
        "pacer_h1": {"result": 88, "process": 30, "total": 118, "kuadran": "KUADRAN 1"},
        "kecukupan_lop_juli": {"target_scal_rkap": "2,10M", "kebutuhan_lop": "4,20M", "est_rev_lop": "6,05M", "ach": 144},
        "visit_juli": {"target_cm": 16, "jml_visit_cm": 18, "ach_cm": 113,
                       "target_ytd": 112, "jml_visit_ytd": 120, "ach_ytd": 107},
        "scal_juli": {"scal_bc": "0,12M", "net_scaling": "0,08M"},
        "rincian_juli": {"AO": "0,06M", "MO+": "0,03M", "TERMIN": "0,00M", "RO": "0,01M",
                         "SO": "0,00M", "DO": "-0,01M", "MO-": "0,00M", "ADJ": "-0,01M", "TOTAL": "0,08M"},
        "kecukupan_agustus": {"est_rev_f3f4": "0,55M", "ach_f3f4": 22, "gap_f3f4": "1,95M",
                               "est_rev_all": "6,40M", "ach_all": 151, "gap_all": "-2,10M"},
        "visit_agustus": {"visit_cm": 5, "target_cm": 16, "ach_cm": 31,
                           "visit_ytd": 120, "target_ytd": 128, "ach_ytd": 94},
        "list_cc": [
            {"cc": "Bank Mega Tbk", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Sinar Mas Land", "lop": "ADA LOP", "scal": "ADA SCALING"},
            {"cc": "Summarecon Agung", "lop": "ADA LOP", "scal": "TANPA SCALING"},
        ],
    },
    "Bagus Santoso": {
        "rev_h1": {"real": "9,10M", "ach": 89},
        "lop_visit_h1": {"kecukupan_all_lop": 80, "jml_visit": 70, "target_visit": 96, "ach_visit": 73},
        "scal_h1": {"scal_bc": "2,40M", "net_scaling": "1,05M"},
        "rincian_h1": {"AO": "1,60M", "MO+": "0,05M", "TERMIN": "0,00M", "RO": "0,05M",
                       "SO": "-0,05M", "DO": "-0,50M", "MO-": "-0,04M", "ADJ": "-0,06M", "TOTAL": "1,05M"},
        "pacer_h1": {"result": 60, "process": 18, "total": 78, "kuadran": "KUADRAN 3"},
        "kecukupan_lop_juli": {"target_scal_rkap": "1,40M", "kebutuhan_lop": "2,80M", "est_rev_lop": "3,10M", "ach": 111},
        "visit_juli": {"target_cm": 16, "jml_visit_cm": 12, "ach_cm": 75,
                       "target_ytd": 112, "jml_visit_ytd": 92, "ach_ytd": 82},
        "scal_juli": {"scal_bc": "0,02M", "net_scaling": "-0,06M"},
        "rincian_juli": {"AO": "0,01M", "MO+": "0,00M", "TERMIN": "0,00M", "RO": "0,00M",
                         "SO": "0,00M", "DO": "-0,04M", "MO-": "0,00M", "ADJ": "-0,03M", "TOTAL": "-0,06M"},
        "kecukupan_agustus": {"est_rev_f3f4": "0,15M", "ach_f3f4": 5, "gap_f3f4": "2,85M",
                               "est_rev_all": "3,00M", "ach_all": 96, "gap_all": "0,10M"},
        "visit_agustus": {"visit_cm": 0, "target_cm": 16, "ach_cm": 0,
                           "visit_ytd": 92, "target_ytd": 128, "ach_ytd": 72},
        "list_cc": [
            {"cc": "Pertamina Retail", "lop": "ADA LOP", "scal": "TANPA SCALING"},
            {"cc": "Krakatau Steel", "lop": "TANPA LOP", "scal": "TANPA SCALING"},
        ],
    },
}


SHEET_MAP = {
    "PRS_Bulanan": "prs_monthly",
    "PRS_KPI": "prs_kpi",
    "PRS_Portofolio": "prs_portfolio",
    "PRS_Regional": "prs_regional",
    "AM_Performance": "am",
    "Action_Plan": "action",
}

# 4 sheet Excel khusus untuk detail scorecard AM Performance (satu paket, dibaca bersamaan)
AM_DETAIL_SHEETS = ["AM_Scorecard_H1", "AM_Scorecard_Juli", "AM_Scorecard_Agustus", "AM_List_CC"]


def am_detail_to_sheets(am_detail):
    """dict AM_DETAIL -> 4 DataFrame (untuk didownload sebagai template).
    Kolom ACH & GAP TIDAK disertakan -> sudah dihitung otomatis oleh sistem, tidak perlu diisi manual."""
    h1_rows, juli_rows, agustus_rows, cc_rows = [], [], [], []
    for name, d in am_detail.items():
        h1_rows.append({
            "Name": name, "Real": d["rev_h1"]["real"], "ACH": d["rev_h1"]["ach"],  # ACH di sini masih manual (lihat catatan di app)
            "KecukupanAllLOP": d["lop_visit_h1"]["kecukupan_all_lop"],
            "TargetVisit": d["lop_visit_h1"]["target_visit"], "JmlVisit": d["lop_visit_h1"]["jml_visit"],
            "ScalBC": d["scal_h1"]["scal_bc"], "NetScaling": d["scal_h1"]["net_scaling"],
            "AO": d["rincian_h1"]["AO"], "MOPlus": d["rincian_h1"]["MO+"], "TERMIN": d["rincian_h1"]["TERMIN"],
            "RO": d["rincian_h1"]["RO"], "SO": d["rincian_h1"]["SO"], "DO": d["rincian_h1"]["DO"],
            "MOMinus": d["rincian_h1"]["MO-"], "ADJ": d["rincian_h1"]["ADJ"], "TOTAL": d["rincian_h1"]["TOTAL"],
            "PacerResult": d["pacer_h1"]["result"], "PacerProcess": d["pacer_h1"]["process"],
            "PacerTotal": d["pacer_h1"]["total"], "Kuadran": d["pacer_h1"]["kuadran"],
        })
        juli_rows.append({
            "Name": name,
            "TargetScalRKAP": d["kecukupan_lop_juli"]["target_scal_rkap"],
            "KebutuhanLOP": d["kecukupan_lop_juli"]["kebutuhan_lop"],
            "EstRevLOP": d["kecukupan_lop_juli"]["est_rev_lop"],
            "TargetCM": d["visit_juli"]["target_cm"], "JmlVisitCM": d["visit_juli"]["jml_visit_cm"],
            "TargetYTD": d["visit_juli"]["target_ytd"], "JmlVisitYTD": d["visit_juli"]["jml_visit_ytd"],
            "ScalBC": d["scal_juli"]["scal_bc"], "NetScaling": d["scal_juli"]["net_scaling"],
            "AO": d["rincian_juli"]["AO"], "MOPlus": d["rincian_juli"]["MO+"], "TERMIN": d["rincian_juli"]["TERMIN"],
            "RO": d["rincian_juli"]["RO"], "SO": d["rincian_juli"]["SO"], "DO": d["rincian_juli"]["DO"],
            "MOMinus": d["rincian_juli"]["MO-"], "ADJ": d["rincian_juli"]["ADJ"], "TOTAL": d["rincian_juli"]["TOTAL"],
        })
        agustus_rows.append({
            "Name": name,
            "EstRevF3F4": d["kecukupan_agustus"]["est_rev_f3f4"], "EstRevAll": d["kecukupan_agustus"]["est_rev_all"],
            "VisitCM": d["visit_agustus"]["visit_cm"], "TargetCM": d["visit_agustus"]["target_cm"],
            "VisitYTD": d["visit_agustus"]["visit_ytd"], "TargetYTD": d["visit_agustus"]["target_ytd"],
        })
        for c in d["list_cc"]:
            cc_rows.append({"Name": name, "CC": c["cc"], "KetLOP": c["lop"], "KetScal": c["scal"]})
    return (pd.DataFrame(h1_rows), pd.DataFrame(juli_rows), pd.DataFrame(agustus_rows), pd.DataFrame(cc_rows))


def build_am_detail_from_sheets(sheets):
    """4 sheet Excel -> dict AM_DETAIL. Return None kalau sheet-nya tidak lengkap (fallback ke default).
    ACH & GAP dihitung otomatis di halaman AM Performance, jadi TIDAK dibaca dari sini."""
    if not all(n in sheets and not sheets[n].empty for n in AM_DETAIL_SHEETS):
        return None
    h1 = sheets["AM_Scorecard_H1"].set_index("Name")
    juli = sheets["AM_Scorecard_Juli"].set_index("Name")
    agustus = sheets["AM_Scorecard_Agustus"].set_index("Name")
    cc = sheets["AM_List_CC"]

    result = {}
    for name in h1.index:
        r = h1.loc[name]
        j = juli.loc[name] if name in juli.index else None
        a = agustus.loc[name] if name in agustus.index else None
        cc_rows = cc[cc["Name"] == name]
        result[name] = {
            "rev_h1": {"real": r["Real"], "ach": r["ACH"]},
            "lop_visit_h1": {"kecukupan_all_lop": r["KecukupanAllLOP"],
                              "jml_visit": r["JmlVisit"], "target_visit": r["TargetVisit"]},
            "scal_h1": {"scal_bc": r["ScalBC"], "net_scaling": r["NetScaling"]},
            "rincian_h1": {"AO": r["AO"], "MO+": r["MOPlus"], "TERMIN": r["TERMIN"], "RO": r["RO"],
                           "SO": r["SO"], "DO": r["DO"], "MO-": r["MOMinus"], "ADJ": r["ADJ"], "TOTAL": r["TOTAL"]},
            "pacer_h1": {"result": r["PacerResult"], "process": r["PacerProcess"],
                         "total": r["PacerTotal"], "kuadran": r["Kuadran"]},
            "kecukupan_lop_juli": ({} if j is None else {
                "target_scal_rkap": j["TargetScalRKAP"], "kebutuhan_lop": j["KebutuhanLOP"],
                "est_rev_lop": j["EstRevLOP"]}),
            "visit_juli": ({} if j is None else {
                "target_cm": j["TargetCM"], "jml_visit_cm": j["JmlVisitCM"],
                "target_ytd": j["TargetYTD"], "jml_visit_ytd": j["JmlVisitYTD"]}),
            "scal_juli": ({} if j is None else {"scal_bc": j["ScalBC"], "net_scaling": j["NetScaling"]}),
            "rincian_juli": ({} if j is None else {
                "AO": j["AO"], "MO+": j["MOPlus"], "TERMIN": j["TERMIN"], "RO": j["RO"],
                "SO": j["SO"], "DO": j["DO"], "MO-": j["MOMinus"], "ADJ": j["ADJ"], "TOTAL": j["TOTAL"]}),
            "kecukupan_agustus": ({} if a is None else {
                "est_rev_f3f4": a["EstRevF3F4"], "est_rev_all": a["EstRevAll"]}),
            "visit_agustus": ({} if a is None else {
                "visit_cm": a["VisitCM"], "target_cm": a["TargetCM"],
                "visit_ytd": a["VisitYTD"], "target_ytd": a["TargetYTD"]}),
            "list_cc": [{"cc": row["CC"], "lop": row["KetLOP"], "scal": row["KetScal"]}
                        for _, row in cc_rows.iterrows()],
        }
    return result


SAVED_DATA_PATH = "saved_dashboard_data.xlsx"  # disimpan di server, dipakai bersama semua device

# Taruh file logo di folder "assets/" pada repo GitHub kamu.
# "width" & "height" = ukuran kotak logo dalam pixel, atur sendiri per logo.
# object-fit:contain (di bawah) memastikan logo tidak gepeng/melar walau
# rasio width:height beda dari rasio asli gambarnya.
LOGO_DIR = "assets"
LOGOS = [
    {"file": "logo_prs.png",       "width": 64, "height": 46},
    {"file": "logo_divisi.png",    "width": 60, "height": 50},
    {"file": "logo_telkom.png",    "width": 100, "height": 40},
    {"file": "logo_danantara.png", "width": 90, "height": 46},
]


def render_logos():
    existing = [l for l in LOGOS if os.path.exists(os.path.join(LOGO_DIR, l["file"]))]
    if not existing:
        return
    imgs_html = ""
    for logo in existing:
        path = os.path.join(LOGO_DIR, logo["file"])
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        ext = path.rsplit(".", 1)[-1]
        w, h = logo["width"], logo["height"]
        min_w = max(20, w * 0.55)
        min_h = max(18, h * 0.55)
        imgs_html += f'''<img src="data:image/{ext};base64,{b64}"
            style="width:clamp({min_w}px, {w/6}vw, {w}px);
                   height:clamp({min_h}px, {h/6}vw, {h}px);
                   object-fit:contain;">'''
    st.markdown(f"""
    <div style="display:flex; flex-wrap:wrap; align-items:center; justify-content:center;
                width:100%; gap:22px; margin-bottom:14px;">
        {imgs_html}
    </div>
    """, unsafe_allow_html=True)


def _read_sheets_into(data, is_real, file_like, source_label):
    try:
        sheets = pd.read_excel(file_like, sheet_name=None)
        for sheet_name, key in SHEET_MAP.items():
            if sheet_name in sheets and not sheets[sheet_name].empty:
                data[key] = sheets[sheet_name]
                is_real[key] = True
        am_detail = build_am_detail_from_sheets(sheets)
        if am_detail:
            data["am_detail"] = am_detail
            is_real["am_detail"] = True
        return True
    except Exception as e:
        st.error(f"Gagal membaca {source_label}: {e}")
        return False


def load_data(uploaded_file):
    """
    Prioritas data:
    1. File yang baru diupload di sesi INI -> dipakai, sekaligus disimpan ke disk server
       supaya device lain (HP, dsb) yang buka dashboard yang sama ikut memakainya.
    2. Kalau tidak ada upload baru, tapi ada data tersimpan dari upload SEBELUMNYA
       (dari device manapun) -> otomatis dipakai.
    3. Kalau belum pernah ada upload sama sekali -> pakai data contoh (dummy).
    """
    data = dummy_data()
    data["am_detail"] = AM_DETAIL_DEFAULT
    is_real = {k: False for k in data}
    source = "dummy"

    if uploaded_file is not None:
        # simpan ke disk server supaya persist lintas device
        with open(SAVED_DATA_PATH, "wb") as f:
            f.write(uploaded_file.getbuffer())
        ok = _read_sheets_into(data, is_real, uploaded_file, "file yang baru diupload")
        if ok:
            source = "fresh_upload"
    elif os.path.exists(SAVED_DATA_PATH):
        ok = _read_sheets_into(data, is_real, SAVED_DATA_PATH, "data tersimpan sebelumnya")
        if ok:
            source = "saved_on_server"

    return data, is_real, source


def make_template_excel():
    buf = BytesIO()
    d = dummy_data()
    h1_df, juli_df, agustus_df, cc_df = am_detail_to_sheets(AM_DETAIL_DEFAULT)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        d["prs_monthly"].to_excel(writer, sheet_name="PRS_Bulanan", index=False)
        d["prs_kpi"].to_excel(writer, sheet_name="PRS_KPI", index=False)
        d["prs_portfolio"].to_excel(writer, sheet_name="PRS_Portofolio", index=False)
        d["prs_regional"].to_excel(writer, sheet_name="PRS_Regional", index=False)
        d["am"].to_excel(writer, sheet_name="AM_Performance", index=False)
        d["action"].to_excel(writer, sheet_name="Action_Plan", index=False)
        h1_df.to_excel(writer, sheet_name="AM_Scorecard_H1", index=False)
        juli_df.to_excel(writer, sheet_name="AM_Scorecard_Juli", index=False)
        agustus_df.to_excel(writer, sheet_name="AM_Scorecard_Agustus", index=False)
        cc_df.to_excel(writer, sheet_name="AM_List_CC", index=False)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# TOP-RIGHT ICONS — toggle tema (Gelap/Terang) + panel data (ikon kecil, bukan sidebar)
# ----------------------------------------------------------------------------
_spacer, _theme_col, _icon_col = st.columns([18, 1, 1])
with _theme_col:
    _icon = "☀️" if st.session_state.theme == "dark" else "🌙"
    if st.button(_icon, help="Ganti tampilan Terang/Gelap"):
        st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"
        st.rerun()
with _icon_col:
    with st.popover("⚙️"):
        st.markdown("#### 📁 Data Dashboard")
        uploaded = st.file_uploader("Upload Excel data real (.xlsx)", type=["xlsx"])
        st.download_button(
            "⬇️ Download template Excel",
            data=make_template_excel(),
            file_name="template_dashboard_prs.xlsx",
            help="Isi template ini dengan data real, lalu upload kembali di atas.",
        )
        with st.expander("Format sheet yang dibaca"):
            st.write("""
            - **PRS_Bulanan**: Month, Realisasi, Target
            - **PRS_KPI**: Metric, Value, Target
            - **PRS_Portofolio**: Type, Value
            - **PRS_Regional**: Region, Value
            - **AM_Performance**: Name, Region, Target, Score, Komunikasi, Negosiasi, Kepatuhan, KepuasanKlien, Pelaporan
            - **Action_Plan**: AM, Item, Due, Priority, Status
            - **AM_Scorecard_H1**, **AM_Scorecard_Juli**, **AM_Scorecard_Agustus**, **AM_List_CC**:
              detail scorecard per-AM (4 sheet ini harus lengkap semua supaya terbaca — kalau salah satu
              kosong, tampilan AM Performance tetap pakai data contoh)
            """)

        data, is_real, source = load_data(uploaded)
        n_real = sum(is_real.values())
        n_total = len(is_real)
        if source == "fresh_upload":
            st.success(f"✔ {n_real}/{n_total} sheet tersimpan & aktif untuk SEMUA device yang buka dashboard ini.")
        elif source == "saved_on_server":
            st.info(f"📌 Memakai data tersimpan dari upload sebelumnya ({n_real}/{n_total} sheet real).")
        else:
            st.info("Belum ada data yang pernah diupload — menampilkan data contoh.")

        if os.path.exists(SAVED_DATA_PATH):
            if st.button("🗑️ Reset ke data contoh (hapus data tersimpan)"):
                os.remove(SAVED_DATA_PATH)
                st.rerun()

if "page" not in st.session_state:
    st.session_state.page = "overview"


def goto(p):
    st.session_state.page = p


# ----------------------------------------------------------------------------
# OVERVIEW PAGE
# ----------------------------------------------------------------------------
def render_overview():
    render_logos()
    st.markdown(f"""
    <div style='text-align:center;padding-top:10px;'>
        <div class='pill-title'>DASHBOARD OVERVIEW</div>
        <div class='subtitle'>A monthly performance overview to guide strategy and highlight progress</div>
    </div>
    """, unsafe_allow_html=True)
    st.write("")

    c1, c2, c3 = st.columns(3)
    cards = [
        (c1, "📊", "PRS PERFORMANCE", "Ringkasan kinerja properti & resource", "prs"),
        (c2, "👥", "AM PERFORMANCE", "Kinerja Account Manager per wilayah", "am"),
        (c3, "📱", "ACTION PLAN AM", "Daftar tindak lanjut & status progres", "action"),
    ]
    for col, icon, title, desc, key in cards:
        with col:
            st.markdown(f"""
            <div class="card-box">
                <div class="icon">{icon}</div>
                <h3>{title}</h3>
                <p>{desc}</p>
            </div>
            """, unsafe_allow_html=True)
            st.button(f"Buka {title.title()}", key=f"btn_{key}", use_container_width=True,
                      on_click=goto, args=(key,))


# ----------------------------------------------------------------------------
# PRS PERFORMANCE PAGE
# ----------------------------------------------------------------------------
def render_prs():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown(f"## PRS <span style='color:{CYAN}'>Performance</span>", unsafe_allow_html=True)

    kpi = data["prs_kpi"]
    cols = st.columns(len(kpi))
    for col, (_, row) in zip(cols, kpi.iterrows()):
        delta = row["Value"] - row["Target"]
        col.metric(row["Metric"], f"{row['Value']:.0f}%", f"{delta:+.1f} vs target")

    c1, c2 = st.columns([1.3, 1])
    with c1:
        st.markdown("#### Revenue Achievement Bulanan")
        m = data["prs_monthly"]
        fig = go.Figure()
        fig.add_bar(x=m["Month"], y=m["Realisasi"], name="Realisasi", marker_color=GOLD)
        fig.add_bar(x=m["Month"], y=m["Target"], name="Target", marker_color="rgba(63,214,240,0.4)")
        fig.update_layout(**PLOTLY_LAYOUT, barmode="group", yaxis_ticksuffix="%")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("#### Komposisi Portofolio Aset")
        p = data["prs_portfolio"]
        fig = px.pie(p, names="Type", values="Value", hole=0.55,
                      color_discrete_sequence=[GOLD, CYAN, GREEN, "#8a6ff0"])
        fig.update_layout(**PLOTLY_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Kinerja per Regional")
    r = data["prs_regional"].sort_values("Value", ascending=True)
    fig = go.Figure(go.Bar(x=r["Value"], y=r["Region"], orientation="h", marker_color=GOLD))
    fig.update_layout(**PLOTLY_LAYOUT, xaxis_ticksuffix="%", height=280)
    st.plotly_chart(fig, use_container_width=True)


# ----------------------------------------------------------------------------
# AM PERFORMANCE PAGE
# ----------------------------------------------------------------------------
def _row(label, value, cls=""):
    return f"<div class='am-row {cls}'><span>{label}</span><span>{value}</span></div>"


def _two_col(l1, v1, c1, l2, v2, c2):
    return f"""<div style="display:flex;">
        <div style="flex:1;">{_row(l1, v1, c1)}</div>
        <div style="flex:1;">{_row(l2, v2, c2)}</div>
    </div>"""


def _panel(header_text, sections, gold=False):
    hcls = "am-header gold" if gold else "am-header"
    scls = "am-subheader gold" if gold else "am-subheader"
    inner = ""
    for subheader, body in sections:
        if subheader:
            inner += f"<div class='{scls}'>{subheader}</div>"
        inner += body
    return f"<div class='am-panel'><div class='{hcls}'>{header_text}</div>{inner}</div>"


def _ach(v):
    return "g" if v >= 100 else "r"


def _parse_m(v):
    """Ubah angka format 'X,XXM' (koma = desimal ala Indonesia) jadi float. '5,21M' -> 5.21 ; 88 -> 88.0"""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper().replace("M", "").replace(",", ".")
    return float(s) if s else 0.0


def _fmt_m(v):
    """5.21 -> '5,21M' — dipakai untuk menampilkan hasil hitungan dalam format yang sama dengan input."""
    return f"{v:.2f}".replace(".", ",") + "M"


def _row_ach(label, value_pct):
    """Baris ACH dengan badge warna mencolok (hijau/merah + glow), beda dari baris biasa."""
    cls = _ach(value_pct)
    arrow = "▲" if cls == "g" else "▼"
    return f"<div class='am-row'><span>{label}</span><span class='am-badge {cls}'>{arrow} {value_pct:.0f}%</span></div>"


def _row_gap(label, value_str):
    """Baris GAP dengan badge merah mencolok (lower better -> selalu ditonjolkan sebagai perhatian)."""
    return f"<div class='am-row'><span>{label}</span><span class='am-badge r'>{value_str}</span></div>"


def _two_col_ach(l1, v1, l2, v2):
    return f"""<div style="display:flex;">
        <div style="flex:1;">{_row_ach(l1, v1)}</div>
        <div style="flex:1;">{_row_ach(l2, v2)}</div>
    </div>"""


def _two_col_gap(l1, v1, l2, v2):
    return f"""<div style="display:flex;">
        <div style="flex:1;">{_row_gap(l1, v1)}</div>
        <div style="flex:1;">{_row_gap(l2, v2)}</div>
    </div>"""


def render_am():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown(f"## AM <span style='color:{CYAN}'>Performance</span>", unsafe_allow_html=True)

    names = list(data["am_detail"].keys())
    selected = st.selectbox("Pilih Account Manager", names, label_visibility="collapsed")
    d = data["am_detail"][selected]

    st.markdown(f"<div class='am-banner'>{selected.upper()}</div>", unsafe_allow_html=True)

    col_l, col_r = st.columns([1, 1.6])

    # ---------------- LEFT: Performance H1 ----------------
    with col_l:
        st.markdown(_panel("REV H1", [(None,
            _row("REAL", d["rev_h1"]["real"]) + _row_ach("ACH", d["rev_h1"]["ach"])
        )]), unsafe_allow_html=True)

        lv = d["lop_visit_h1"]
        ach_visit = lv["jml_visit"] / lv["target_visit"] * 100 if lv["target_visit"] else 0  # RUMUS: Jml Visit / Target Visit
        body = (_row("KECUKUPAN ALL LOP", f"{lv['kecukupan_all_lop']}%", _ach(lv["kecukupan_all_lop"]))
                + _row(f"JML VISIT (T. {lv['target_visit']} Visit)", lv["jml_visit"])
                + _row_ach("ACH VISIT", ach_visit))
        st.markdown(_panel("LOP & VISIT H1", [(None, body)]), unsafe_allow_html=True)

        sc = d["scal_h1"]
        body = _row("SCAL BC (AO, MO+, TERMIN, RO)", sc["scal_bc"]) + _row("NET SCALING", sc["net_scaling"])
        st.markdown(_panel("SCAL H1", [(None, body)], gold=True), unsafe_allow_html=True)

        rn = d["rincian_h1"]
        body = ("".join(_row(k, rn[k], "g") for k in ["AO", "MO+", "TERMIN", "RO"])
                + "".join(_row(k, rn[k], "r") for k in ["SO", "DO", "MO-", "ADJ"])
                + _row("TOTAL", rn["TOTAL"], "total"))
        st.markdown(_panel("RINCIAN NET SCALING H1", [(None, body)], gold=True), unsafe_allow_html=True)

        pc = d["pacer_h1"]
        body = (_row("TOTAL POIN RESULT", pc["result"], "g")
                + _row("TOTAL POIN PROCESS", pc["process"], "r")
                + _row("TOTAL ALL", pc["total"], "total")
                + _row("KUADRAN AM", pc["kuadran"], "total"))
        st.markdown(_panel("PACER H1", [(None, body)]), unsafe_allow_html=True)

    # ---------------- RIGHT: Juli & Agustus ----------------
    with col_r:
        r1, r2 = st.columns(2)
        with r1:
            k = d["kecukupan_lop_juli"]
            est_rev_lop = _parse_m(k["est_rev_lop"])
            kebutuhan_lop = _parse_m(k["kebutuhan_lop"])
            ach_lop = est_rev_lop / kebutuhan_lop * 100 if kebutuhan_lop else 0  # RUMUS: Est Rev LOP / Kebutuhan LOP
            body1 = (_row("TARGET SCAL RKAP", k["target_scal_rkap"])
                     + _row("KEBUTUHAN LOP (2X T.SCAL)", k["kebutuhan_lop"])
                     + _row("EST REV LOP", k["est_rev_lop"])
                     + _row_ach("ACH", ach_lop))
            v = d["visit_juli"]
            ach_cm = v["jml_visit_cm"] / v["target_cm"] * 100 if v["target_cm"] else 0      # RUMUS: Jml Visit CM / Target CM
            ach_ytd = v["jml_visit_ytd"] / v["target_ytd"] * 100 if v["target_ytd"] else 0  # RUMUS: Jml Visit YTD / Target YTD
            body2 = (_row("TARGET CM", v["target_cm"])
                     + _row("JML VISIT CM", v["jml_visit_cm"])
                     + _row_ach("ACH CM", ach_cm)
                     + _row("TARGET YTD", v["target_ytd"])
                     + _row("JML VISIT YTD", v["jml_visit_ytd"])
                     + _row_ach("ACH YTD", ach_ytd))
            st.markdown(_panel("LOP & VISIT JULI (Cut off 3 Agustus)", [
                ("KECUKUPAN ALL LOP", body1), ("VISIT JULI & YTD JULI", body2),
            ]), unsafe_allow_html=True)

        with r2:
            sj = d["scal_juli"]
            body1 = _row("SCAL BC (AO, MO+, TERMIN, RO)", sj["scal_bc"]) + _row("NET SCALING", sj["net_scaling"])
            rj = d["rincian_juli"]
            body2 = ("".join(_row(k, rj[k], "g") for k in ["AO", "MO+", "TERMIN", "RO"])
                     + "".join(_row(k, rj[k], "r") for k in ["SO", "DO", "MO-", "ADJ"])
                     + _row("TOTAL", rj["TOTAL"], "total"))
            st.markdown(_panel("SCALING JULI 2026", [
                ("SCAL JULI", body1), ("RINCIAN NET SCALING JULI", body2),
            ], gold=True), unsafe_allow_html=True)

        r3, r4 = st.columns(2)
        with r3:
            ka = d["kecukupan_agustus"]
            est_rev_f3f4 = _parse_m(ka["est_rev_f3f4"])
            est_rev_all = _parse_m(ka["est_rev_all"])
            # RUMUS: dibandingkan terhadap Kebutuhan LOP Juli (sumber kebutuhan yang sama)
            ach_f3f4 = est_rev_f3f4 / kebutuhan_lop * 100 if kebutuhan_lop else 0
            ach_all = est_rev_all / kebutuhan_lop * 100 if kebutuhan_lop else 0
            gap_f3f4 = kebutuhan_lop - est_rev_f3f4   # RUMUS: Kebutuhan LOP - Est Rev (Lower Better)
            gap_all = kebutuhan_lop - est_rev_all
            body = (_two_col("EST REV LOP F3-F4", ka["est_rev_f3f4"], "", "EST REV ALL LOP", ka["est_rev_all"], "")
                    + _two_col_ach("ACH", ach_f3f4, "ACH", ach_all)
                    + _two_col_gap("GAP (Lower Better)", _fmt_m(gap_f3f4), "GAP (Lower Better)", _fmt_m(gap_all)))
            st.markdown(_panel("KECUKUPAN LOP AGUSTUS", [(None, body)]), unsafe_allow_html=True)

        with r4:
            va = d["visit_agustus"]
            ach_cm_ag = va["visit_cm"] / va["target_cm"] * 100 if va["target_cm"] else 0      # RUMUS: Visit CM / Target CM
            ach_ytd_ag = va["visit_ytd"] / va["target_ytd"] * 100 if va["target_ytd"] else 0  # RUMUS: Visit YTD / Target YTD
            body = (_two_col(f"VISIT CM (T. {va['target_cm']} Visit)", va["visit_cm"], "",
                              f"VISIT YTD (T. {va['target_ytd']} Visit)", va["visit_ytd"], "")
                    + _two_col_ach("ACH", ach_cm_ag, "ACH", ach_ytd_ag))
            st.markdown(_panel("VISIT AGUSTUS", [(None, body)], gold=True), unsafe_allow_html=True)

        st.markdown(_panel("LIST CC AM (Cut off 3 Agustus)", [(None, "")]), unsafe_allow_html=True)
        cc_df = pd.DataFrame([
            {"AM 2026": selected, "CC": c["cc"], "KET LOP 2026": c["lop"], "KET SCAL BC YTD JULI": c["scal"]}
            for c in d["list_cc"]
        ])
        st.dataframe(cc_df, use_container_width=True, hide_index=True)


# ----------------------------------------------------------------------------
# ACTION PLAN AM PAGE
# ----------------------------------------------------------------------------
def render_action():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown(f"## Action Plan <span style='color:{CYAN}'>AM</span>", unsafe_allow_html=True)

    a = data["action"]
    filt = st.radio("Filter", ["Semua", "Prioritas Tinggi", "Berjalan", "Selesai"],
                     horizontal=True, label_visibility="collapsed")
    view = a.copy()
    if filt == "Prioritas Tinggi":
        view = view[view["Priority"] == "Tinggi"]
    elif filt == "Berjalan":
        view = view[view["Status"] == "Berjalan"]
    elif filt == "Selesai":
        view = view[view["Status"] == "Selesai"]

    def color_priority(v):
        c = {"Tinggi": RED, "Sedang": GOLD, "Rendah": GREEN}.get(v, MUTED)
        return f"color:{c}; font-weight:700;"

    def color_status(v):
        c = {"Selesai": GREEN, "Berjalan": CYAN, "Belum Mulai": MUTED}.get(v, MUTED)
        return f"color:{c}; font-weight:700;"

    styled = view.style.map(color_priority, subset=["Priority"]).map(color_status, subset=["Status"])
    st.dataframe(styled, use_container_width=True, hide_index=True)

    st.caption(f"Menampilkan {len(view)} dari {len(a)} action item.")


# ----------------------------------------------------------------------------
# ROUTER
# ----------------------------------------------------------------------------
page = st.session_state.page
if page == "overview":
    render_overview()
elif page == "prs":
    render_prs()
elif page == "am":
    render_am()
elif page == "action":
    render_action()

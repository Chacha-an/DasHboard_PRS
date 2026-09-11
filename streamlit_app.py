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
.am-header.gold{{ background:{CYAN}; }}
.am-subheader{{ background:{GLOW1}; color:{CYAN}; font-weight:700; padding:6px 14px; text-align:center; font-size:11.5px;}}
.am-subheader.gold{{ background:{GLOW1}; color:{CYAN}; }}
.am-row{{ display:flex; justify-content:space-between; padding:7px 14px; font-size:12.5px; border-bottom:1px solid {ROWBORDER};}}
.am-row span:first-child{{ color:{MUTED}; }}
.am-row span:last-child{{ font-weight:700; color:{TEXT}; }}
.am-row.g{{ background:rgba(15,157,88,0.12); }}
.am-row.g span:last-child{{ color:{GREEN}; }}
.am-row.r{{ background:rgba(217,48,37,0.12); }}
.am-row.r span:last-child{{ color:{RED}; }}
.am-row.total{{ background:{TOTALBG}; font-weight:800; }}
.am-row.total span{{ color:{TEXT}; }}
.am-row.label{{ justify-content:center; }}
.am-row.label span{{ width:100%; text-align:center; }}
.am-row.label.hl{{ background:{CYAN}; }}
.am-row.label.hl span{{ color:{HEADER_TEXT} !important; font-weight:800; letter-spacing:0.4px; }}

/* Badge ACH & GAP — lebih mencolok dari baris biasa */
.am-badge{{
    display:inline-block; padding:3px 12px; border-radius:999px;
    font-weight:900; font-size:12.5px; letter-spacing:0.3px;
}}
.am-badge.g{{ background:{GREEN}; color:#ffffff; box-shadow:0 0 12px {GREEN_GLOW}; }}
.am-badge.r{{ background:{RED}; color:#ffffff; box-shadow:0 0 12px {RED_GLOW}; }}
.am-badge.n{{ background:{CYAN}; color:{HEADER_TEXT}; box-shadow:0 0 12px {GLOW1}; }}
.am-bignum{{
    display:inline-block; padding:3px 12px; border-radius:999px; font-weight:900; font-size:12.5px; letter-spacing:0.3px;
}}
.am-bignum.g{{ background:{GREEN}; color:#ffffff; box-shadow:0 0 14px {GREEN_GLOW}; }}
.am-bignum.r{{ background:{RED}; color:#ffffff; box-shadow:0 0 14px {RED_GLOW}; }}
.am-photo-frame{{
    width:100%; aspect-ratio:4/5; max-height:320px; min-height:150px;
    border-radius:16px; overflow:hidden;
    border:1px solid {BORDER}; box-shadow:0 0 20px {GLOW1}; background:{CARD};
    display:flex; align-items:center; justify-content:center;
}}
.am-photo-frame img{{ width:100%; height:100%; object-fit:cover; object-position:top; }}

/* Stat box bold + center — pengganti st.metric() yang suka kepotong */
.am-stat{{
    text-align:center; padding:12px 6px; border:1px solid {BORDER}; border-radius:12px;
    background:{CARD}; margin-bottom:8px;
}}
.am-stat .lbl{{ font-size:10px; font-weight:800; color:{MUTED}; text-transform:uppercase; letter-spacing:0.3px; margin-bottom:5px; }}
.am-stat .val{{ font-size:22px; font-weight:900; color:{TEXT}; }}

/* Sub-judul section — kotak solid orange + highlight, dipakai sebagai pemisah antar baris */
.am-section-title{{
    background:{GOLD}; color:{HEADER_TEXT}; font-weight:900; font-size:14px; text-align:center;
    letter-spacing:0.6px; padding:11px 16px; border-radius:10px; margin:20px 0 12px;
    box-shadow:0 0 16px {GLOW2};
}}

/* Grid kuadran PACER — 2x2, foto+nama AM ditempatkan sesuai kuadrannya */
.quad-grid{{
    display:grid; grid-template-columns:1fr 1fr; grid-template-rows:1fr 1fr;
    gap:8px; padding:4px;
}}
.quad-box{{
    border:1px dashed {BORDER}; border-radius:10px; padding:8px; min-height:150px;
    display:flex; flex-direction:column; align-items:center;
}}
.quad-box .quad-label{{
    font-weight:800; font-size:10.5px; color:{MUTED}; text-align:center; margin-bottom:6px;
    text-transform:uppercase; letter-spacing:0.4px;
}}
.quad-chips{{ display:flex; flex-wrap:wrap; justify-content:center; gap:4px; width:100%; }}
.quad-chip{{ display:flex; flex-direction:column; align-items:center; width:56px; text-align:center; }}
.quad-chip img, .quad-chip .noimg{{
    width:40px; height:40px; border-radius:50%; object-fit:cover; border:2px solid {CYAN};
}}
.quad-chip .noimg{{ background:{BG2}; display:flex; align-items:center; justify-content:center; font-size:18px; }}
.quad-chip .nm{{ font-size:8.5px; font-weight:700; margin-top:3px; line-height:1.15; color:{TEXT}; }}
.quad-empty{{ color:{MUTED}; font-size:10px; text-align:center; padding-top:20px; }}

/* Tabel HTML custom (LOP AM 2026, Visit 2026) — dipakai supaya header pasti berwarna,
   karena st.dataframe kadang tidak nurut styling header dari pandas */
.html-table{{ width:100%; border-collapse:collapse; font-size:12px; white-space:nowrap; }}
.html-table th{{ padding:9px 10px; text-align:center; font-weight:800; }}
.html-table td{{ padding:7px 10px; text-align:center; border-bottom:1px solid {ROWBORDER}; color:{TEXT}; }}
.html-table tbody tr:nth-child(even) td{{ background:rgba(120,120,120,0.05); }}
.html-table td.g{{ color:{GREEN}; font-weight:700; }}
.html-table td.r{{ color:{RED}; font-weight:700; }}
.html-table tr.total td{{ background:{TOTALBG}; font-weight:800; }}
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
    prs_monthly_revenue = pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"],
        "Target": [14000000000, 14000000000, 15000000000, 15000000000, 16000000000, 16000000000,
                   17000000000, 17000000000, 18000000000, 18000000000, 19000000000, 20000000000],
        "Realisasi": [13200000000, 14500000000, 14800000000, 15600000000, 15900000000, 17100000000,
                      16500000000, 18000000000, 17600000000, 19200000000, 18800000000, 20500000000],
    })
    prs_monthly_ngtma = pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"],
        "Target": [800000000, 800000000, 900000000, 900000000, 1000000000, 1000000000,
                   1000000000, 1100000000, 1100000000, 1200000000, 1200000000, 1300000000],
        "Realisasi": [500000000, 600000000, 700000000, 650000000, 800000000, 750000000,
                      900000000, 950000000, 1000000000, 1050000000, 1100000000, 1150000000],
    })
    prs_detail_revenue = pd.DataFrame({
        "Metric": ["NON POTS", "POTS", "IFRS"],
        "Value": [120000000000, 68000000000, -8000000000],
    })
    prs_detail_mitra = pd.DataFrame({
        "Mitra": ["Telkomsel", "Telkom Akses", "Mitra Distribusi A", "Mitra Distribusi B"],
        "Nominal": [45000000000, 30000000000, 12000000000, 8000000000],
    })
    return {"prs_monthly_revenue": prs_monthly_revenue, "prs_monthly_ngtma": prs_monthly_ngtma,
            "prs_detail_revenue": prs_detail_revenue, "prs_detail_mitra": prs_detail_mitra}


# ----------------------------------------------------------------------------
# AM PERFORMANCE — detail scorecard per Account Manager (dummy, 3 contoh).
# Untuk tambah AM lain, copy salah satu blok ini dan ganti angkanya.
# ----------------------------------------------------------------------------
AM_DETAIL_DEFAULT = {
    "Mar'atus Sholicha": {
        "photo": "photo_maratus_sholicha.png",
        "display_name": "Licha",  # contoh nama panggilan custom — muncul di kuadran, boleh dikosongkan
        "period_ytd": "AGUSTUS 2026", "period_month": "SEPTEMBER", "cutoff_date": "19 September 2026",
        "real_rev": {
            "cm": {"target": 1730000000, "real": 1490000000},
            "ytd": {"target": 13480000000, "real": 11790000000},
            "ngtma": {"real": 10000000, "ach": 1},
            "detail": {"non_pots": 7480000000, "pots": 3210000000, "ifrs": -390000000, "total": 10300000000},
        },
        "real_scaling": {
            "cm": {"scal_bc": 10000000, "net_scaling": 0},
            "ytd": {"scal_bc": 800000000, "net_scaling": 200000000},
            "detail": {"AO": 340000000, "MO+": 440000000, "TERMIN": 0, "RO": 20000000,
                       "SO": -20000000, "DO": -60000000, "MO-": -20000000, "ADJ": -510000000, "TOTAL": 200000000},
        },
        "pacer_juli": {"result": 70, "ach_revenue": 86, "ach_scaling": 64, "win_rate": 160, "kecukupan_qualified": 42,
                       "process": 33, "kec_lop": 228, "jml_visit": 17, "target_visit": 16,
                       "total_pacer": 103, "kuadran": "KUADRAN 3"},
        "pacer_ytd": {"result": 62, "ach_revenue": 91, "ach_scaling": 25, "win_rate": 160, "kecukupan_qualified": 42,
                      "process": 34, "kec_lop": 228, "jml_visit": 110, "target_visit": 96,
                      "total_pacer": 96, "kuadran": "KUADRAN 3"},
        "kecukupan_lop": {"target": 4860000000, "est_rev_f3f4": 760000000, "est_rev_all": 15270000000},
        "lop_fy2026": {"target_scal_rkap": 2380000000, "est_rev_f0f2": 14500000000, "est_rev_f3f4": 760000000},
        "visit_monthly": {"target": 144, "months": {"Jan": 14, "Feb": 14, "Mar": 13, "Apr": 21, "May": 11,
                                                       "Jun": 20, "Jul": 17, "Aug": 27, "Sept": 4}},
        "visit_bulan": {"cm": {"jml_visit": 7, "target": 16}, "ytd": {"jml_visit": 154, "target": 128}},
        "list_cc": [
            {"cc": "PT Bara Tabang", "jml_lop": "1 LOP", "jml_scal": "TIDAK ADA SCALING"},
            {"cc": "Podomoro Group", "jml_lop": "10 LOP", "jml_scal": "0,16 M"},
            {"cc": "Donggi Senoro LNG", "jml_lop": "14 LOP", "jml_scal": "0,11 M"},
            {"cc": "Cifor", "jml_lop": "3 LOP", "jml_scal": "0 M"},
            {"cc": "Roheda Sejati / Plaza Oleos", "jml_lop": "3 LOP", "jml_scal": "TIDAK ADA SCALING"},
            {"cc": "Bayan Resources", "jml_lop": "7 LOP", "jml_scal": "0,01 M"},
            {"cc": "Artha Telekomindo", "jml_lop": "76 LOP", "jml_scal": "0,1 M"},
            {"cc": "Orica", "jml_lop": "TIDAK ADA LOP", "jml_scal": "0,83 M"},
            {"cc": "Permata Senayan (Arthatel)", "jml_lop": "TIDAK ADA LOP", "jml_scal": "TIDAK ADA SCALING"},
        ],
        "list_lop": [
            {"lop_id": "P26-200854", "proj": "APL - PT GPS Astinet dan Indibiz", "est_bc": 0.0045, "ket_lob": "NEW LOB"},
            {"lop_id": "P26-204525", "proj": "PSB Metro P2MP 200 Mbps PT LSAG Cable Indonesia, Artha Industrial Hill Blok E Kav 20-21", "est_bc": 0.010095, "ket_lob": "NEW LOB"},
            {"lop_id": "P26-204355", "proj": "PSB Astinet Fit 50 Mbps PT. MAP ZONA ADIPERKASA (DIGIPLUS) Bassura City Mall, Jl. Jend", "est_bc": 0.00491, "ket_lob": "NEW LOB"},
            {"lop_id": "P26-204145", "proj": "PSB Metro P2MP 5 Mbps PT. Toray Industries Indonesia Cab Semarang, Jl. Raya Tegalpana", "est_bc": 0.00484, "ket_lob": "NEW LOB"},
            {"lop_id": "P26-203327", "proj": "PSB Astinet Dedicated 30 Mbps Erablue Electronic 8 Lokasi (2 Quote) Blok 001 Sawah Kali", "est_bc": 0.0047141, "ket_lob": "CO Aug"},
            {"lop_id": "P26-203556", "proj": "PSB Astinet Fit 50 Mbps PT. MAP ZONA ADIPERKASA (DIGIMAP) Ponorogo City Center, Jl.", "est_bc": 0.003, "ket_lob": "NEW LOB"},
        ],
    },
    "Rina Wulandari": {
        "photo": "photo_rina_wulandari.png",
        "period_ytd": "AGUSTUS 2026", "period_month": "SEPTEMBER", "cutoff_date": "19 September 2026",
        "real_rev": {
            "cm": {"target": 1900000000, "real": 2050000000},
            "ytd": {"target": 15200000000, "real": 16800000000},
            "ngtma": {"real": 30000000, "ach": 3},
            "detail": {"non_pots": 9600000000, "pots": 5100000000, "ifrs": 2100000000, "total": 16800000000},
        },
        "real_scaling": {
            "cm": {"scal_bc": 50000000, "net_scaling": 40000000},
            "ytd": {"scal_bc": 1200000000, "net_scaling": 950000000},
            "detail": {"AO": 600000000, "MO+": 300000000, "TERMIN": 50000000, "RO": 60000000,
                       "SO": -10000000, "DO": -30000000, "MO-": -10000000, "ADJ": -10000000, "TOTAL": 950000000},
        },
        "pacer_juli": {"result": 92, "ach_revenue": 108, "ach_scaling": 95, "win_rate": 175, "kecukupan_qualified": 60,
                       "process": 55, "kec_lop": 260, "jml_visit": 19, "target_visit": 16,
                       "total_pacer": 147, "kuadran": "KUADRAN 1"},
        "pacer_ytd": {"result": 96, "ach_revenue": 111, "ach_scaling": 79, "win_rate": 170, "kecukupan_qualified": 58,
                      "process": 60, "kec_lop": 250, "jml_visit": 130, "target_visit": 96,
                      "total_pacer": 156, "kuadran": "KUADRAN 1"},
        "kecukupan_lop": {"target": 5200000000, "est_rev_f3f4": 2100000000, "est_rev_all": 19800000000},
        "lop_fy2026": {"target_scal_rkap": 2100000000, "est_rev_f0f2": 5200000000, "est_rev_f3f4": 2100000000},
        "visit_monthly": {"target": 144, "months": {"Jan": 18, "Feb": 20, "Mar": "Cuti Melahirkan", "Apr": "Cuti Melahirkan",
                                                       "May": 12, "Jun": 22, "Jul": 24, "Aug": 26, "Sept": 8}},
        "visit_bulan": {"cm": {"jml_visit": 18, "target": 16}, "ytd": {"jml_visit": 138, "target": 128}},
        "list_cc": [
            {"cc": "Bank Mega Tbk", "jml_lop": "6 LOP", "jml_scal": "0,45 M"},
            {"cc": "Sinar Mas Land", "jml_lop": "9 LOP", "jml_scal": "0,30 M"},
            {"cc": "Summarecon Agung", "jml_lop": "4 LOP", "jml_scal": "TIDAK ADA SCALING"},
        ],
        "list_lop": [
            {"lop_id": "P26-210044", "proj": "APL - PT Bank Mega Cabang Sudirman", "est_bc": 0.0061, "ket_lob": "NEW LOB"},
            {"lop_id": "P26-211320", "proj": "PSB Metro P2MP 500 Mbps Sinar Mas Land Plaza BSD", "est_bc": 0.0125, "ket_lob": "NEW LOB"},
        ],
    },
    "Bagus Santoso": {
        "photo": "photo_bagus_santoso.png",
        "period_ytd": "AGUSTUS 2026", "period_month": "SEPTEMBER", "cutoff_date": "19 September 2026",
        "real_rev": {
            "cm": {"target": 1200000000, "real": 780000000},
            "ytd": {"target": 9600000000, "real": 6700000000},
            "ngtma": {"real": 0, "ach": 0},
            "detail": {"non_pots": 4800000000, "pots": 2400000000, "ifrs": -500000000, "total": 6700000000},
        },
        "real_scaling": {
            "cm": {"scal_bc": 5000000, "net_scaling": -5000000},
            "ytd": {"scal_bc": 300000000, "net_scaling": 60000000},
            "detail": {"AO": 120000000, "MO+": 40000000, "TERMIN": 0, "RO": 10000000,
                       "SO": -20000000, "DO": -60000000, "MO-": -10000000, "ADJ": -20000000, "TOTAL": 60000000},
        },
        "pacer_juli": {"result": 48, "ach_revenue": 65, "ach_scaling": 40, "win_rate": 120, "kecukupan_qualified": 25,
                       "process": 20, "kec_lop": 150, "jml_visit": 10, "target_visit": 16,
                       "total_pacer": 68, "kuadran": "KUADRAN 4"},
        "pacer_ytd": {"result": 52, "ach_revenue": 70, "ach_scaling": 20, "win_rate": 130, "kecukupan_qualified": 30,
                      "process": 22, "kec_lop": 140, "jml_visit": 78, "target_visit": 96,
                      "total_pacer": 74, "kuadran": "KUADRAN 4"},
        "kecukupan_lop": {"target": 3400000000, "est_rev_f3f4": 200000000, "est_rev_all": 2900000000},
        "lop_fy2026": {"target_scal_rkap": 1700000000, "est_rev_f0f2": 1200000000, "est_rev_f3f4": 200000000},
        "visit_monthly": {"target": 144, "months": {"Jan": 6, "Feb": 8, "Mar": 10, "Apr": 9, "May": 7,
                                                       "Jun": 12, "Jul": 14, "Aug": 15, "Sept": 5}},
        "visit_bulan": {"cm": {"jml_visit": 4, "target": 16}, "ytd": {"jml_visit": 78, "target": 128}},
        "list_cc": [
            {"cc": "Pertamina Retail", "jml_lop": "2 LOP", "jml_scal": "TIDAK ADA SCALING"},
            {"cc": "Krakatau Steel", "jml_lop": "TIDAK ADA LOP", "jml_scal": "TIDAK ADA SCALING"},
        ],
        "list_lop": [
            {"lop_id": "P26-198740", "proj": "PSB Astinet Fit 20 Mbps Pertamina Retail SPBU Cilandak", "est_bc": 0.0035, "ket_lob": "NEW LOB"},
        ],
    },
}


# Urutan bulan penuh 1 tahun — dipakai untuk deteksi kolom bulan otomatis di sheet AM_Visit_Monthly
# (tabel Visit 2026 hanya menampilkan bulan yang benar-benar ada kolomnya di Excel, urut kalender)
FULL_MONTH_ORDER = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"]

SHEET_MAP = {
    "PRS_Monthly_Revenue": "prs_monthly_revenue",
    "PRS_Monthly_NGTMA": "prs_monthly_ngtma",
    "PRS_Detail_Revenue": "prs_detail_revenue",
    "PRS_Detail_Mitra": "prs_detail_mitra",
}

# 7 sheet Excel khusus untuk detail scorecard AM Performance (satu paket, dibaca bersamaan)
AM_DETAIL_SHEETS = ["AM_Summary", "AM_RealRev_Scaling", "AM_Pacer", "AM_Kecukupan_LOP_Visit",
                    "AM_List_CC", "AM_List_LOP"]


def am_detail_to_sheets(am_detail):
    """dict AM_DETAIL -> 8 DataFrame (untuk didownload sebagai template).
    ACH & GAP TIDAK disertakan -> dihitung otomatis, tidak perlu diisi manual."""
    summary_rows, rev_rows, pacer_rows, lop_rows, cc_rows, lop_id_rows, lop_fy_rows, visit_m_rows = [], [], [], [], [], [], [], []
    for name, d in am_detail.items():
        summary_rows.append({
            "Name": name, "PhotoFile": d.get("photo", ""), "DisplayName": d.get("display_name", ""),
            "PeriodYTD": d["period_ytd"], "PeriodMonth": d["period_month"], "CutoffDate": d["cutoff_date"],
        })
        rr, rs = d["real_rev"], d["real_scaling"]
        rev_rows.append({
            "Name": name,
            "RevCMTarget": rr["cm"]["target"], "RevCMReal": rr["cm"]["real"],
            "RevYTDTarget": rr["ytd"]["target"], "RevYTDReal": rr["ytd"]["real"],
            "NGTMAReal": rr["ngtma"]["real"], "NGTMAAch": rr["ngtma"]["ach"],
            "NonPOTS": rr["detail"]["non_pots"], "POTS": rr["detail"]["pots"],
            "IFRS": rr["detail"]["ifrs"], "RevTotal": rr["detail"]["total"],
            "ScalCMBC": rs["cm"]["scal_bc"], "ScalCMNet": rs["cm"]["net_scaling"],
            "ScalYTDBC": rs["ytd"]["scal_bc"], "ScalYTDNet": rs["ytd"]["net_scaling"],
            "AO": rs["detail"]["AO"], "MOPlus": rs["detail"]["MO+"], "TERMIN": rs["detail"]["TERMIN"],
            "RO": rs["detail"]["RO"], "SO": rs["detail"]["SO"], "DO": rs["detail"]["DO"],
            "MOMinus": rs["detail"]["MO-"], "ADJ": rs["detail"]["ADJ"], "ScalTOTAL": rs["detail"]["TOTAL"],
        })
        for period_label, p in [("Juli", d["pacer_juli"]), ("YTD Juli", d["pacer_ytd"])]:
            pacer_rows.append({
                "Name": name, "Period": period_label, "Result": p["result"],
                "AchRevenue": p["ach_revenue"], "AchScaling": p["ach_scaling"], "WinRate": p["win_rate"],
                "KecukupanQualified": p["kecukupan_qualified"], "Process": p["process"], "KecLOP": p["kec_lop"],
                "JmlVisit": p["jml_visit"], "TargetVisit": p["target_visit"],
                "TotalPacer": p["total_pacer"], "Kuadran": p["kuadran"],
            })
        kl, vb = d["kecukupan_lop"], d["visit_bulan"]
        lop_rows.append({
            "Name": name, "TargetKecLOP": kl["target"], "EstRevF3F4": kl["est_rev_f3f4"], "EstRevAll": kl["est_rev_all"],
            "VisitCMJml": vb["cm"]["jml_visit"], "VisitCMTarget": vb["cm"]["target"],
            "VisitYTDJml": vb["ytd"]["jml_visit"], "VisitYTDTarget": vb["ytd"]["target"],
        })
        for c in d["list_cc"]:
            cc_rows.append({"Name": name, "CC": c["cc"], "JmlLOP": c["jml_lop"], "JmlScalBC": c["jml_scal"]})
        for l in d["list_lop"]:
            lop_id_rows.append({"Name": name, "LopID": l["lop_id"], "Proj": l["proj"],
                                 "EstBC": l["est_bc"], "KetLOB": l["ket_lob"]})
        lf = d.get("lop_fy2026")
        if lf:
            lop_fy_rows.append({"Name": name, "TargetScalRKAP": lf["target_scal_rkap"],
                                 "EstRevF0F2": lf["est_rev_f0f2"], "EstRevF3F4": lf["est_rev_f3f4"]})
        vm = d.get("visit_monthly")
        if vm:
            row = {"Name": name, "Target": vm["target"]}
            row.update(vm["months"])
            visit_m_rows.append(row)
    return (pd.DataFrame(summary_rows), pd.DataFrame(rev_rows), pd.DataFrame(pacer_rows),
            pd.DataFrame(lop_rows), pd.DataFrame(cc_rows), pd.DataFrame(lop_id_rows),
            pd.DataFrame(lop_fy_rows), pd.DataFrame(visit_m_rows))


def build_am_detail_from_sheets(sheets):
    """6 sheet Excel -> dict AM_DETAIL. Return None kalau sheet-nya tidak lengkap (fallback ke default).
    ACH & GAP dihitung otomatis di halaman AM Performance, jadi TIDAK dibaca dari sini."""
    if not all(n in sheets and not sheets[n].empty for n in AM_DETAIL_SHEETS):
        return None
    summary = sheets["AM_Summary"].set_index("Name")
    rev = sheets["AM_RealRev_Scaling"].set_index("Name")
    pacer = sheets["AM_Pacer"]
    lop = sheets["AM_Kecukupan_LOP_Visit"].set_index("Name")
    cc = sheets["AM_List_CC"]
    lop_id = sheets["AM_List_LOP"]
    lop_fy = sheets.get("AM_LOP_FY2026")  # opsional — sheet baru, tidak wajib ada di file lama
    if lop_fy is not None and not lop_fy.empty:
        lop_fy = lop_fy.set_index("Name")
    visit_m = sheets.get("AM_Visit_Monthly")  # opsional
    if visit_m is not None and not visit_m.empty:
        visit_m = visit_m.set_index("Name")
        visit_months_detected = [m for m in FULL_MONTH_ORDER if m in visit_m.columns]  # otomatis ikut kolom yang ADA di Excel
    else:
        visit_months_detected = []

    result = {}
    for name in summary.index:
        s = summary.loc[name]
        r = rev.loc[name] if name in rev.index else None
        l = lop.loc[name] if name in lop.index else None
        p_rows = pacer[pacer["Name"] == name]
        p_juli = p_rows[p_rows["Period"] == "Juli"]
        p_ytd = p_rows[p_rows["Period"] == "YTD Juli"]

        def _pacer_dict(row):
            return {"result": row["Result"], "ach_revenue": row["AchRevenue"], "ach_scaling": row["AchScaling"],
                    "win_rate": row["WinRate"], "kecukupan_qualified": row["KecukupanQualified"],
                    "process": row["Process"], "kec_lop": row["KecLOP"], "jml_visit": row["JmlVisit"],
                    "target_visit": row["TargetVisit"], "total_pacer": row["TotalPacer"], "kuadran": row["Kuadran"]}

        cc_rows = cc[cc["Name"] == name]
        lop_id_rows = lop_id[lop_id["Name"] == name]

        result[name] = {
            "photo": s.get("PhotoFile", ""),
            "display_name": s.get("DisplayName", "") if pd.notna(s.get("DisplayName", "")) else "",
            "period_ytd": s["PeriodYTD"], "period_month": s["PeriodMonth"], "cutoff_date": s["CutoffDate"],
            "real_rev": ({} if r is None else {
                "cm": {"target": r["RevCMTarget"], "real": r["RevCMReal"]},
                "ytd": {"target": r["RevYTDTarget"], "real": r["RevYTDReal"]},
                "ngtma": {"real": r["NGTMAReal"], "ach": r["NGTMAAch"]},
                "detail": {"non_pots": r["NonPOTS"], "pots": r["POTS"], "ifrs": r["IFRS"], "total": r["RevTotal"]},
            }),
            "real_scaling": ({} if r is None else {
                "cm": {"scal_bc": r["ScalCMBC"], "net_scaling": r["ScalCMNet"]},
                "ytd": {"scal_bc": r["ScalYTDBC"], "net_scaling": r["ScalYTDNet"]},
                "detail": {"AO": r["AO"], "MO+": r["MOPlus"], "TERMIN": r["TERMIN"], "RO": r["RO"],
                           "SO": r["SO"], "DO": r["DO"], "MO-": r["MOMinus"], "ADJ": r["ADJ"], "TOTAL": r["ScalTOTAL"]},
            }),
            "pacer_juli": (_pacer_dict(p_juli.iloc[0]) if len(p_juli) else {}),
            "pacer_ytd": (_pacer_dict(p_ytd.iloc[0]) if len(p_ytd) else {}),
            "kecukupan_lop": ({} if l is None else {
                "target": l["TargetKecLOP"], "est_rev_f3f4": l["EstRevF3F4"], "est_rev_all": l["EstRevAll"]}),
            "visit_bulan": ({} if l is None else {
                "cm": {"jml_visit": l["VisitCMJml"], "target": l["VisitCMTarget"]},
                "ytd": {"jml_visit": l["VisitYTDJml"], "target": l["VisitYTDTarget"]},
            }),
            "list_cc": [{"cc": row["CC"], "jml_lop": row["JmlLOP"], "jml_scal": row["JmlScalBC"]}
                        for _, row in cc_rows.iterrows()],
            "list_lop": [{"lop_id": row["LopID"], "proj": row["Proj"], "est_bc": row["EstBC"], "ket_lob": row["KetLOB"]}
                         for _, row in lop_id_rows.iterrows()],
            "lop_fy2026": ({} if lop_fy is None or name not in lop_fy.index else {
                "target_scal_rkap": lop_fy.loc[name, "TargetScalRKAP"],
                "est_rev_f0f2": lop_fy.loc[name, "EstRevF0F2"],
                "est_rev_f3f4": lop_fy.loc[name, "EstRevF3F4"]}),
            "visit_monthly": ({} if visit_m is None or name not in visit_m.index else {
                "target": visit_m.loc[name, "Target"],
                "months": {m: visit_m.loc[name, m] for m in visit_months_detected}}),
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


@st.cache_data(show_spinner=False)
def _b64_file(path, mtime):
    """Baca file & encode ke base64 SEKALI saja per file (di-cache Streamlit).
    'mtime' disertakan supaya cache otomatis refresh kalau filenya diganti/redeploy."""
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def render_logos():
    existing = [l for l in LOGOS if os.path.exists(os.path.join(LOGO_DIR, l["file"]))]
    if not existing:
        return
    imgs_html = ""
    for logo in existing:
        path = os.path.join(LOGO_DIR, logo["file"])
        b64 = _b64_file(path, os.path.getmtime(path))
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


# Taruh foto AM di folder "assets/photos/" pada repo GitHub, nama file sesuai kolom
# PhotoFile di sheet AM_Summary. Kalau file belum ada, otomatis tampil ikon placeholder.
PHOTO_DIR = "assets/photos"


def render_am_photo(filename):
    path = os.path.join(PHOTO_DIR, filename) if filename else None
    if path and os.path.exists(path):
        b64 = _b64_file(path, os.path.getmtime(path))
        ext = path.rsplit(".", 1)[-1]
        inner = f'<img src="data:image/{ext};base64,{b64}">'
    else:
        inner = f"<span style='font-size:56px; color:{MUTED};'>👤</span>"
    st.markdown(f"<div class='am-photo-frame'>{inner}</div>", unsafe_allow_html=True)


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
    summary_df, rev_df, pacer_df, lop_df, cc_df, lop_id_df, lop_fy_df, visit_m_df = am_detail_to_sheets(AM_DETAIL_DEFAULT)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        d["prs_monthly_revenue"].to_excel(writer, sheet_name="PRS_Monthly_Revenue", index=False)
        d["prs_monthly_ngtma"].to_excel(writer, sheet_name="PRS_Monthly_NGTMA", index=False)
        d["prs_detail_revenue"].to_excel(writer, sheet_name="PRS_Detail_Revenue", index=False)
        d["prs_detail_mitra"].to_excel(writer, sheet_name="PRS_Detail_Mitra", index=False)
        summary_df.to_excel(writer, sheet_name="AM_Summary", index=False)
        rev_df.to_excel(writer, sheet_name="AM_RealRev_Scaling", index=False)
        pacer_df.to_excel(writer, sheet_name="AM_Pacer", index=False)
        lop_df.to_excel(writer, sheet_name="AM_Kecukupan_LOP_Visit", index=False)
        cc_df.to_excel(writer, sheet_name="AM_List_CC", index=False)
        lop_id_df.to_excel(writer, sheet_name="AM_List_LOP", index=False)
        lop_fy_df.to_excel(writer, sheet_name="AM_LOP_FY2026", index=False)
        visit_m_df.to_excel(writer, sheet_name="AM_Visit_Monthly", index=False)
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
            - **PRS_Monthly_Revenue**: Month, Target, Realisasi (12 baris, Jan-Des, angka penuh)
            - **PRS_Monthly_NGTMA**: Month, Target, Realisasi (12 baris, Jan-Des, angka penuh)
            - **PRS_Detail_Revenue**: Metric (NON POTS/POTS/IFRS), Value (angka penuh)
            - **PRS_Detail_Mitra**: Mitra, Nominal (angka penuh, baris sebanyak jumlah mitra)
            - **AM_Summary**, **AM_RealRev_Scaling**, **AM_Pacer**, **AM_Kecukupan_LOP_Visit**,
              **AM_List_CC**, **AM_List_LOP**: detail scorecard per-AM (6 sheet ini harus lengkap semua
              supaya terbaca — kalau salah satu kosong, tampilan AM Performance tetap pakai data contoh).
              Kolom `PhotoFile` di **AM_Summary** diisi nama file foto yang ditaruh di folder `assets/photos/`.
              Kolom `DisplayName` di **AM_Summary** (opsional) untuk nama panggilan custom yang tampil di kuadran PACER —
              kosongkan kalau mau pakai format otomatis (Nama Depan + Inisial Belakang).
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

    c1, c2 = st.columns(2)
    cards = [
        (c1, "📊", "PRS PERFORMANCE", "Ringkasan kinerja properti & resource", "prs"),
        (c2, "👥", "AM PERFORMANCE", "Kinerja Account Manager per wilayah", "am"),
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
def render_quadrant_chart(title, period_key, gold=False):
    """Grid 2x2: Kuadran 1 kanan-atas, 2 kiri-atas, 3 kanan-bawah, 4 kiri-bawah.
    Isinya foto+nama SEMUA AM, dikelompokkan berdasarkan field 'kuadran' di data masing-masing."""
    groups = {"KUADRAN 1": [], "KUADRAN 2": [], "KUADRAN 3": [], "KUADRAN 4": []}
    for name, d in data["am_detail"].items():
        k = str(d.get(period_key, {}).get("kuadran", "")).upper().strip()
        if k in groups:
            groups[k].append((name, d.get("photo"), d.get("display_name")))

    def _chips(members):
        if not members:
            return "<div class='quad-empty'>Belum ada AM</div>"
        html = "<div class='quad-chips'>"
        for name, photo, display_name in members:
            path = os.path.join(PHOTO_DIR, photo) if photo else None
            if path and os.path.exists(path):
                b64 = _b64_file(path, os.path.getmtime(path))
                ext = path.rsplit(".", 1)[-1]
                img_html = f"<img src='data:image/{ext};base64,{b64}'>"
            else:
                img_html = "<div class='noimg'>👤</div>"
            if display_name:
                short = display_name                      # PRIORITAS: nama custom dari Excel kalau diisi
            else:
                parts = name.split()
                short = f"{parts[0]} {parts[-1][0]}." if len(parts) > 1 else parts[0]  # fallback otomatis
            html += f"<div class='quad-chip'>{img_html}<div class='nm'>{short}</div></div>"
        return html + "</div>"

    grid_html = f"""<div class='quad-grid'>
        <div class='quad-box'><div class='quad-label'>KUADRAN 2</div>{_chips(groups["KUADRAN 2"])}</div>
        <div class='quad-box'><div class='quad-label'>KUADRAN 1</div>{_chips(groups["KUADRAN 1"])}</div>
        <div class='quad-box'><div class='quad-label'>KUADRAN 4</div>{_chips(groups["KUADRAN 4"])}</div>
        <div class='quad-box'><div class='quad-label'>KUADRAN 3</div>{_chips(groups["KUADRAN 3"])}</div>
    </div>"""
    st.markdown(_panel(title, [(None, grid_html)], gold=gold), unsafe_allow_html=True)


def _monthly_bar_chart(title, monthly_df):
    """Bar chart Target vs Realisasi per bulan, dengan badge ACH% di atas tiap pasang bar."""
    months = monthly_df["Month"].tolist()
    targets_m = [t / 1_000_000_000 for t in monthly_df["Target"]]
    reals_m = [r / 1_000_000_000 for r in monthly_df["Realisasi"]]
    achs = [(r / t * 100 if t else 0) for r, t in zip(reals_m, targets_m)]

    fig = go.Figure()
    fig.add_bar(x=months, y=targets_m, name="Target", marker_color="rgba(63,214,240,0.4)")
    fig.add_bar(x=months, y=reals_m, name="Realisasi", marker_color=GOLD)

    annotations = []
    for mo, ach, r, t in zip(months, achs, reals_m, targets_m):
        color = GREEN if ach >= 100 else RED
        annotations.append(dict(
            x=mo, y=max(r, t) * 1.15, text=f"<b>{ach:.0f}%</b>", showarrow=False,
            font=dict(size=10, color="#ffffff"),
            bgcolor=color, borderpad=3,
        ))
    fig.update_layout(**PLOTLY_LAYOUT, barmode="group", annotations=annotations,
                       yaxis_title="Miliar (M)")
    st.markdown(f"##### {title}")
    st.plotly_chart(fig, use_container_width=True)


def render_prs():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown(f"## PRS <span style='color:{CYAN}'>Performance</span>", unsafe_allow_html=True)

    # ============== SECTION 1: PACER AM PRS ==============
    render_section_title("PACER AM PRS")
    q1, q2 = st.columns(2)
    with q1:
        render_quadrant_chart("KUADRAN CM", "pacer_juli")
    with q2:
        render_quadrant_chart("KUADRAN YTD", "pacer_ytd", gold=True)

    # ============== SECTION 2: PERFORMANCE REVENUE ==============
    render_section_title("PERFORMANCE REVENUE")
    r1, r2 = st.columns(2)
    with r1:
        _monthly_bar_chart("Revenue 2026", data["prs_monthly_revenue"])
        dr = data["prs_detail_revenue"]
        total_rev = dr["Value"].sum()
        body = "".join(_row(row["Metric"], _fmt_m(row["Value"])) for _, row in dr.iterrows())
        body += _row("TOTAL", _fmt_m(total_rev), "total")
        st.markdown(_panel("DETAIL REVENUE", [(None, body)]), unsafe_allow_html=True)

    with r2:
        _monthly_bar_chart("NGTMA 2026", data["prs_monthly_ngtma"])
        dm = data["prs_detail_mitra"]
        total_mitra = dm["Nominal"].sum()
        body2 = "".join(_row(row["Mitra"], _fmt_m(row["Nominal"])) for _, row in dm.iterrows())
        body2 += _row("TOTAL", _fmt_m(total_mitra), "total")
        st.markdown(_panel("DETAIL MITRA", [(None, body2)], gold=True), unsafe_allow_html=True)

    # ============== SECTION 3: VISIT DAN LOP AM 2026 ==============
    render_section_title("VISIT DAN LOP AM 2026")

    st.markdown(_panel("LOP AM 2026", [(None, "")]), unsafe_allow_html=True)
    lop_rows = []
    for name, d in data["am_detail"].items():
        lf = d.get("lop_fy2026")
        if not lf:
            continue
        target_scal = lf["target_scal_rkap"] / 1_000_000_000
        kebutuhan_lop = target_scal * 2                                    # RUMUS: 2 x Target Scal RKAP
        est_f0f2 = lf["est_rev_f0f2"] / 1_000_000_000
        est_f3f4 = lf["est_rev_f3f4"] / 1_000_000_000
        est_all = est_f0f2 + est_f3f4                                       # RUMUS: F0-F2 + F3-F4
        pct_f3f4 = est_f3f4 / kebutuhan_lop * 100 if kebutuhan_lop else 0    # RUMUS: F3-F4 / Kebutuhan LOP
        pct_all = est_all / kebutuhan_lop * 100 if kebutuhan_lop else 0      # RUMUS: All / Kebutuhan LOP
        gap_all = kebutuhan_lop - est_all                                   # RUMUS: Kebutuhan LOP - All (Lower Better)
        lop_rows.append({
            "AM 2026": name, "T. SCAL RKAP FY 2026": target_scal, "KEBUTUHAN LOP (2X Target Scaling)": kebutuhan_lop,
            "EST REV LOP F0-F2": est_f0f2, "EST REV LOP F3-F4": est_f3f4, "EST REV ALL LOP": est_all,
            "% KECUKUPAN LOP F3-F4": pct_f3f4, "% KECUKUPAN ALL LOP": pct_all,
            "GAP ALL LOP (Lower Better)": gap_all,
        })

    if lop_rows:
        total_row = {"AM 2026": "TOTAL"}
        for col in ["T. SCAL RKAP FY 2026", "KEBUTUHAN LOP (2X Target Scaling)",
                    "EST REV LOP F0-F2", "EST REV LOP F3-F4", "EST REV ALL LOP"]:
            total_row[col] = sum(r[col] for r in lop_rows)
        tk = total_row["KEBUTUHAN LOP (2X Target Scaling)"]
        total_row["% KECUKUPAN LOP F3-F4"] = total_row["EST REV LOP F3-F4"] / tk * 100 if tk else 0
        total_row["% KECUKUPAN ALL LOP"] = total_row["EST REV ALL LOP"] / tk * 100 if tk else 0
        total_row["GAP ALL LOP (Lower Better)"] = tk - total_row["EST REV ALL LOP"]

        def _lop_row_html(r, is_total=False):
            f34_cls = "g" if r["% KECUKUPAN LOP F3-F4"] >= 100 else "r"
            all_cls = "g" if r["% KECUKUPAN ALL LOP"] >= 100 else "r"
            gap_cls = "g" if r["GAP ALL LOP (Lower Better)"] < 0 else "r"
            tr_cls = "total" if is_total else ""
            return f"""<tr class="{tr_cls}">
                <td style="text-align:left; font-weight:{'800' if is_total else '600'};">{r['AM 2026']}</td>
                <td>{r['T. SCAL RKAP FY 2026']:.2f}M</td>
                <td>{r['KEBUTUHAN LOP (2X Target Scaling)']:.2f}M</td>
                <td>{r['EST REV LOP F0-F2']:.2f}M</td>
                <td>{r['EST REV LOP F3-F4']:.2f}M</td>
                <td>{r['EST REV ALL LOP']:.2f}M</td>
                <td class="{f34_cls}">{r['% KECUKUPAN LOP F3-F4']:.0f}%</td>
                <td class="{all_cls}">{r['% KECUKUPAN ALL LOP']:.0f}%</td>
                <td class="{gap_cls}">{r['GAP ALL LOP (Lower Better)']:.2f}M</td>
            </tr>"""

        body_html = "".join(_lop_row_html(r) for r in lop_rows) + _lop_row_html(total_row, is_total=True)
        headers = ["AM 2026", "T. SCAL RKAP FY 2026", "KEBUTUHAN LOP (2X Target Scaling)",
                   "EST REV LOP F0-F2", "EST REV LOP F3-F4", "EST REV ALL LOP",
                   "% KECUKUPAN LOP F3-F4", "% KECUKUPAN ALL LOP", "GAP ALL LOP (Lower Better)"]
        render_html_table(headers, body_html, CYAN)
    else:
        st.caption("Belum ada data 'lop_fy2026' — isi lewat sheet AM_LOP_FY2026 di Excel.")

    st.write("")
    st.markdown(_panel("VISIT 2026", [(None, "")], gold=True), unsafe_allow_html=True)
    # Deteksi otomatis bulan mana saja yang ADA datanya (union dari semua AM), urut kalender.
    # Jadi kalau Excel baru diisi sampai Agustus, kolom cuma sampai Agustus — nambah otomatis
    # begitu bulan berikutnya diisi di Excel dan diupload ulang.
    months_present = set()
    for _, d in data["am_detail"].items():
        vm = d.get("visit_monthly")
        if vm:
            months_present.update(vm["months"].keys())
    MONTHS_ACTIVE = [m for m in FULL_MONTH_ORDER if m in months_present]

    visit_rows = []
    for name, d in data["am_detail"].items():
        vm = d.get("visit_monthly")
        if not vm:
            continue
        row = {"NAMA AM": name, "TARGET": vm["target"]}
        total = 0
        for mo in MONTHS_ACTIVE:
            val = vm["months"].get(mo, "")
            row[mo] = val
            if isinstance(val, (int, float)):
                total += val                                    # RUMUS: Total = jumlah bulan yang ada datanya
        row["TOTAL"] = total
        row["ACH"] = f"{(total / vm['target'] * 100 if vm['target'] else 0):.0f}%"   # RUMUS: Total / Target
        visit_rows.append(row)

    if visit_rows:
        total_row = {"NAMA AM": "TOTAL VISIT", "TARGET": ""}
        grand_total = 0
        for mo in MONTHS_ACTIVE:
            month_sum = sum(r[mo] for r in visit_rows if isinstance(r[mo], (int, float)))
            total_row[mo] = month_sum
            grand_total += month_sum
        total_row["TOTAL"] = grand_total
        total_row["ACH"] = ""

        def _visit_row_html(r, is_total=False):
            ach_cls = ""
            if isinstance(r.get("ACH"), str) and r["ACH"].endswith("%"):
                try:
                    pct = float(r["ACH"][:-1])
                    ach_cls = "g" if pct >= 100 else "r"
                except ValueError:
                    pass
            month_cells = "".join(f"<td>{r[mo]}</td>" for mo in MONTHS_ACTIVE)
            tr_cls = "total" if is_total else ""
            return f"""<tr class="{tr_cls}">
                <td style="text-align:left; font-weight:{'800' if is_total else '600'};">{r['NAMA AM']}</td>
                <td>{r['TARGET']}</td>
                {month_cells}
                <td style="font-weight:700;">{r['TOTAL']}</td>
                <td class="{ach_cls}">{r.get('ACH', '')}</td>
            </tr>"""

        body_html = "".join(_visit_row_html(r) for r in visit_rows) + _visit_row_html(total_row, is_total=True)
        headers = ["NAMA AM", "TARGET"] + MONTHS_ACTIVE + ["TOTAL", "ACH"]
        render_html_table(headers, body_html, CYAN)
        st.caption("Catatan: kalau ada AM cuti/tidak ada data di bulan tertentu, kolom bulan itu akan kosong "
                   "(bukan sel gabungan bertuliskan status seperti di Excel aslinya).")
    else:
        st.caption("Belum ada data 'visit_monthly' — isi lewat sheet AM_Visit_Monthly di Excel.")


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
    """Selalu mengembalikan angka RUPIAH PENUH (float).
    Menerima angka penuh langsung (mis. 11240000000) — format baru & yang dianjurkan —
    ATAU format lama 'X,XXM' ('5,21M' -> 5.21 miliar) supaya data lama tetap kompatibel."""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper()
    if s.endswith("M"):
        return float(s[:-1].replace(",", ".")) * 1_000_000_000
    return float(s.replace(",", ".")) if s else 0.0


def _fmt_m(v):
    """11240000000 -> '11,24M' (M = Miliar Rupiah) — dipakai untuk menampilkan angka penuh dalam format ringkas."""
    return f"{v/1_000_000_000:.2f}".replace(".", ",") + "M"


def _row_ach(label, value_pct):
    """Baris ACH dengan badge warna mencolok (hijau/merah + glow), beda dari baris biasa."""
    cls = _ach(value_pct)
    arrow = "▲" if cls == "g" else "▼"
    return f"<div class='am-row'><span>{label}</span><span class='am-badge {cls}'>{arrow} {value_pct:.0f}%</span></div>"


def _row_gap(label, raw_value, value_str):
    """Badge GAP: minus (Est Rev sudah lebihi Kebutuhan LOP) = hijau, positif (masih kurang) = merah."""
    cls = "g" if raw_value < 0 else "r"
    return f"<div class='am-row'><span>{label}</span><span class='am-badge {cls}'>{value_str}</span></div>"


def _row_pill(label, value_str, cls="n"):
    """Baris dengan badge mencolok generik (bukan format persen) — untuk Kecukupan All LOP, Kuadran AM, dll."""
    return f"<div class='am-row'><span>{label}</span><span class='am-badge {cls}'>{value_str}</span></div>"


def _row_label(text, highlight=False):
    """Baris label section (teks tunggal, di tengah) — dipakai sebagai pemisah/section title di dalam panel."""
    cls = "label hl" if highlight else "label"
    return f"<div class='am-row {cls}'><span>{text}</span></div>"


def _stat_box(label, value):
    """Kotak angka bold + center — pengganti st.metric() yang suka kepotong labelnya."""
    return f"<div class='am-stat'><div class='lbl'>{label}</div><div class='val'>{value}</div></div>"


def render_section_title(text):
    """Sub-judul section: kotak solid + highlight, dipakai sebagai pemisah antar baris di halaman AM Performance."""
    st.markdown(f"<div class='am-section-title'>{text}</div>", unsafe_allow_html=True)


def render_html_table(headers, body_rows_html, header_bg):
    """Render tabel HTML custom dengan header berwarna solid — dipakai untuk tabel besar
    (LOP AM 2026, Visit 2026) supaya warna header pasti tampil, tidak tergantung st.dataframe."""
    head_html = "".join(f"<th>{h}</th>" for h in headers)
    html = f"""<div style="overflow-x:auto; border:1px solid {BORDER}; border-radius:10px;">
    <table class="html-table">
        <thead><tr style="background:{header_bg}; color:{HEADER_TEXT};">{head_html}</tr></thead>
        <tbody>{body_rows_html}</tbody>
    </table>
    </div>"""
    st.markdown(html, unsafe_allow_html=True)


def _two_col_ach(l1, v1, l2, v2):
    return f"""<div style="display:flex;">
        <div style="flex:1;">{_row_ach(l1, v1)}</div>
        <div style="flex:1;">{_row_ach(l2, v2)}</div>
    </div>"""


def _two_col_gap(l1, raw1, str1, l2, raw2, str2):
    return f"""<div style="display:flex;">
        <div style="flex:1;">{_row_gap(l1, raw1, str1)}</div>
        <div style="flex:1;">{_row_gap(l2, raw2, str2)}</div>
    </div>"""


def _pacer_card(title, p, gold=False):
    """Kartu PACER (Result & Process sebagai badge angka besar), pakai style panel yang sama dengan bagian lain."""
    result_cls = "g" if p["result"] >= 75 else "r"
    process_cls = "g" if p["process"] >= 25 else "r"
    result_row = f"""<div class='am-row'><span>TOTAL POIN RESULT</span>
        <span class='am-bignum {result_cls}'>{p['result']}</span></div>"""
    sub_result = (_row("Ach Revenue", f"{p['ach_revenue']}%")
                  + _row("Ach Scaling", f"{p['ach_scaling']}%")
                  + _row("Win Rate", f"{p['win_rate']}%")
                  + _row("Kecukupan Qualified", f"{p['kecukupan_qualified']}%"))
    process_row = f"""<div class='am-row'><span>TOTAL POIN PROCESS</span>
        <span class='am-bignum {process_cls}'>{p['process']}</span></div>"""
    sub_process = (_row("KEC LOP", f"{p['kec_lop']}%")
                   + _row(f"Jml VISIT (T. {p['target_visit']})", p["jml_visit"]))
    total_row = _row("TOTAL POIN PACER", p["total_pacer"], "total")
    kuadran_row = _row_pill("KUADRAN AM", p["kuadran"], "n")
    return _panel(title, [
        (None, result_row), (None, sub_result),
        (None, process_row), (None, sub_process),
        (None, total_row + kuadran_row),
    ], gold=gold)


def render_am():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown(f"## AM <span style='color:{CYAN}'>Performance</span>", unsafe_allow_html=True)

    names = list(data["am_detail"].keys())
    selected = st.selectbox("Pilih Account Manager", names, label_visibility="collapsed")
    d = data["am_detail"][selected]
    month_ytd = d["period_ytd"].split()[0] if d.get("period_ytd") else ""

    # ============== BARIS 1: Foto+Nama | PACER Juli | PACER YTD Juli ==============
    render_section_title("PACER AM PRS")
    col_photo, col_pacer1, col_pacer2 = st.columns([0.8, 1, 1])

    with col_photo:
        render_am_photo(d.get("photo"))
        st.markdown(f"<div class='am-banner' style='font-size:14px; padding:12px; margin-top:12px;'>{selected.upper()}</div>",
                    unsafe_allow_html=True)

    with col_pacer1:
        st.markdown(_pacer_card("PACER JULI", d["pacer_juli"]), unsafe_allow_html=True)

    with col_pacer2:
        st.markdown(_pacer_card("PACER YTD JULI", d["pacer_ytd"], gold=True), unsafe_allow_html=True)

    # ============== BARIS 2: Real Rev | Detail Rev YTD | Real Scaling | Detail Net Scaling YTD ==============
    render_section_title(f"PERFORMANCE YTD {d['period_ytd']}")
    col_rr, col_drv, col_rs, col_dns = st.columns(4)

    rr = d["real_rev"]
    cm_ach = rr["cm"]["real"] / rr["cm"]["target"] * 100 if rr["cm"]["target"] else 0    # RUMUS: Real / Target
    ytd_ach = rr["ytd"]["real"] / rr["ytd"]["target"] * 100 if rr["ytd"]["target"] else 0

    with col_rr:
        body = (_row_label(f"PERFORMANCE {month_ytd}", highlight=True)
                + _row("CM Target", _fmt_m(rr["cm"]["target"])) + _row("CM Real (Incl IFRS)", _fmt_m(rr["cm"]["real"]))
                + _row_ach("CM ACH", cm_ach)
                + _row_label(f"PERFORMANCE YTD {month_ytd}", highlight=True)
                + _row("YTD Target", _fmt_m(rr["ytd"]["target"])) + _row("YTD Real (Incl IFRS)", _fmt_m(rr["ytd"]["real"]))
                + _row_ach("YTD ACH", ytd_ach)
                + _row("NGTMA Real", _fmt_m(rr["ngtma"]["real"]))
                + _row_pill("NGTMA ACH", f"{rr['ngtma']['ach']}%", _ach(rr["ngtma"]["ach"])))
        st.markdown(_panel("REAL REV", [(None, body)]), unsafe_allow_html=True)

    with col_drv:
        det = rr["detail"]
        body2 = (_row("NON POTS", _fmt_m(det["non_pots"])) + _row("POTS", _fmt_m(det["pots"]))
                 + _row("IFRS", _fmt_m(det["ifrs"])) + _row("TOTAL", _fmt_m(det["total"]), "total"))
        st.markdown(_panel("DETAIL REVENUE YTD", [(None, body2)], gold=True), unsafe_allow_html=True)

    with col_rs:
        rs = d["real_scaling"]
        body3 = (_row_label(f"PERFORMANCE {month_ytd}")
                 + _row("CM Scal BC", _fmt_m(rs["cm"]["scal_bc"])) + _row("CM Net Scaling", _fmt_m(rs["cm"]["net_scaling"]))
                 + _row_label(f"PERFORMANCE YTD {month_ytd}")
                 + _row("YTD Scal BC", _fmt_m(rs["ytd"]["scal_bc"])) + _row("YTD Net Scaling", _fmt_m(rs["ytd"]["net_scaling"])))
        st.markdown(_panel("REAL SCALING", [(None, body3)]), unsafe_allow_html=True)

    with col_dns:
        dn = rs["detail"]
        body4 = ("".join(_row(k, _fmt_m(dn[k]), "g") for k in ["AO", "MO+", "TERMIN", "RO"])
                 + "".join(_row(k, _fmt_m(dn[k]), "r") for k in ["SO", "DO", "MO-", "ADJ"])
                 + _row("TOTAL", _fmt_m(dn["TOTAL"]), "total"))
        st.markdown(_panel("DETAIL NET SCALING YTD", [(None, body4)], gold=True), unsafe_allow_html=True)

    # ============== BARIS BAWAH: Kecukupan LOP & Visit | List CC | List LOP ==============
    render_section_title(f"ENSURING {d['period_month']}")
    col_lop, col_cc, col_lobtable = st.columns([1, 1.3, 1.3])

    with col_lop:
        kl = d["kecukupan_lop"]
        target = kl["target"]
        f34, allv = kl["est_rev_f3f4"], kl["est_rev_all"]
        ach_f34 = f34 / target * 100 if target else 0     # RUMUS: Est Rev / Target Kec. LOP
        ach_all = allv / target * 100 if target else 0
        gap_f34 = target - f34                              # RUMUS: Target - Est Rev (Lower Better)
        gap_all = target - allv
        body_f34 = _row("EST REV 2026", _fmt_m(f34)) + _row_ach("ACH", ach_f34) + _row_gap("GAP (Lower Better)", gap_f34, _fmt_m(gap_f34))
        body_all = _row("EST REV 2026", _fmt_m(allv)) + _row_ach("ACH", ach_all) + _row_gap("GAP (Lower Better)", gap_all, _fmt_m(gap_all))
        st.markdown(_panel(f"KECUKUPAN LOP {d['period_month']}", [
            (None, _row("TARGET KEC. LOP", _fmt_m(target))),
            ("LOP F3-F4", body_f34), ("ALL LOP (F0-F4)", body_all),
        ]), unsafe_allow_html=True)

        vb = d["visit_bulan"]
        cm_v_ach = vb["cm"]["jml_visit"] / vb["cm"]["target"] * 100 if vb["cm"]["target"] else 0   # RUMUS: Jml Visit / Target
        ytd_v_ach = vb["ytd"]["jml_visit"] / vb["ytd"]["target"] * 100 if vb["ytd"]["target"] else 0
        body_cm = _row("JML VISIT", vb["cm"]["jml_visit"]) + _row_ach(f"ACH (T. {vb['cm']['target']} Visit)", cm_v_ach)
        body_ytd = _row("JML VISIT", vb["ytd"]["jml_visit"]) + _row_ach(f"ACH (T. {vb['ytd']['target']} Visit)", ytd_v_ach)
        st.markdown(_panel(f"VISIT {d['period_month']}", [
            ("VISIT CM", body_cm), ("VISIT YTD", body_ytd),
        ], gold=True), unsafe_allow_html=True)

    with col_cc:
        jml_cc = len(d["list_cc"])                                                              # RUMUS: jumlah baris di List CC
        jml_tanpa_lop = sum(1 for c in d["list_cc"] if "TIDAK ADA" in str(c["jml_lop"]).upper())  # RUMUS: hitung yang "TIDAK ADA LOP"
        jml_tanpa_scal = sum(1 for c in d["list_cc"] if "TIDAK ADA" in str(c["jml_scal"]).upper()) # RUMUS: hitung yang "TIDAK ADA SCALING"
        stats_html = f"""<div style="display:flex; gap:8px;">
            <div style="flex:1;">{_stat_box("JML CC", jml_cc)}</div>
            <div style="flex:1;">{_stat_box("TANPA LOP", jml_tanpa_lop)}</div>
            <div style="flex:1;">{_stat_box("TANPA SCAL", jml_tanpa_scal)}</div>
        </div>"""
        st.markdown(stats_html, unsafe_allow_html=True)
        st.markdown(_panel(f"LIST CC (Cut off {d['cutoff_date']})", [(None, "")]), unsafe_allow_html=True)
        cc_df = pd.DataFrame([
            {"CC": c["cc"], "JML LOP 2026": c["jml_lop"], f"JML SCAL BC YTD {d['period_ytd'].split()[0]}": c["jml_scal"]}
            for c in d["list_cc"]
        ])
        st.dataframe(cc_df, use_container_width=True, hide_index=True)

    with col_lobtable:
        jml_lob = len(d["list_lop"])                                    # RUMUS: jumlah baris di List LOP
        est_nilai_bc = sum(l["est_bc"] for l in d["list_lop"])           # RUMUS: jumlah kolom Est BC
        stats_html = f"""<div style="display:flex; gap:8px;">
            <div style="flex:1;">{_stat_box("JML LOB", jml_lob)}</div>
            <div style="flex:1;">{_stat_box("EST NILAI BC", f"{est_nilai_bc:.4f}")}</div>
        </div>"""
        st.markdown(stats_html, unsafe_allow_html=True)
        st.markdown(_panel("LIST LOP ID", [(None, "")]), unsafe_allow_html=True)
        lop_df = pd.DataFrame([
            {"LOP ID": l["lop_id"], "PROJ": l["proj"], "EST BC": l["est_bc"], "KET LOB": l["ket_lob"]}
            for l in d["list_lop"]
        ])
        st.dataframe(lop_df, use_container_width=True, hide_index=True)


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

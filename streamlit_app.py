"""
PRS PriDE - Dashboard Overview (Streamlit version)
Jalankan dengan: streamlit run streamlit_app.py

Struktur:
- Sidebar: upload file Excel (opsional) untuk perbandingan data REAL vs dummy.
- 4 halaman: Overview, PRS Performance, AM Performance, Action Plan AM.
"""

import os
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
# THEME / CSS (teal gelap + glow, senada dengan versi HTML)
# ----------------------------------------------------------------------------
CYAN = "#3fd6f0"
GOLD = "#f5b942"
GREEN = "#4fd18b"
RED = "#f0685f"
BG = "#06222e"
CARD = "#0f3f52"
MUTED = "#9fc4cf"

st.markdown(f"""
<style>
.stApp {{
    background:
      radial-gradient(1200px 600px at 15% -10%, rgba(63,214,240,0.10), transparent 60%),
      radial-gradient(900px 500px at 100% 0%, rgba(245,185,66,0.08), transparent 55%),
      linear-gradient(180deg, {BG}, #0a3244 40%, {BG});
    color: #f5fbfd;
}}
[data-testid="stSidebar"] {{ background:#08283a; border-right:1px solid rgba(140,220,235,0.18); }}
h1,h2,h3 {{ font-family:'Trebuchet MS', sans-serif; font-weight:800; }}

.pill-title{{
    display:inline-block; padding:14px 46px; border-radius:999px;
    border:1.5px solid {CYAN}; background:rgba(63,214,240,0.06);
    box-shadow:0 0 24px rgba(63,214,240,0.35), inset 0 0 18px rgba(63,214,240,0.10);
    font-size:32px; font-weight:900; letter-spacing:2px; margin-bottom:6px;
}}
.subtitle{{ font-style:italic; color:{MUTED}; font-size:16px; margin-bottom:10px;}}

.card-box{{
    background:linear-gradient(160deg, {CARD}, #0a3244);
    border:1px solid rgba(140,220,235,0.18); border-radius:20px; padding:26px 16px;
    text-align:center; height:100%;
}}
.card-box .icon{{ font-size:42px; margin-bottom:10px; }}
.card-box h3{{ font-size:15px; letter-spacing:1px; margin-bottom:6px;}}
.card-box p{{ font-size:12px; color:{MUTED}; }}

.metric-card{{
    background:{CARD}; border:1px solid rgba(140,220,235,0.18); border-radius:14px;
    padding:16px 18px; margin-bottom:6px;
}}
.metric-card .label{{ font-size:11px; color:{MUTED}; text-transform:uppercase; letter-spacing:1px; font-weight:600;}}
.metric-card .value{{ font-size:28px; font-weight:800; margin-top:4px;}}
.metric-card .delta.up{{ color:{GREEN}; font-size:12px; font-weight:600;}}
.metric-card .delta.down{{ color:{RED}; font-size:12px; font-weight:600;}}

div[data-testid="stMetric"]{{
    background:{CARD}; border:1px solid rgba(140,220,235,0.18); border-radius:14px; padding:12px 16px;
}}
.stButton>button{{
    background:{CARD}; color:white; border:1px solid rgba(140,220,235,0.25); border-radius:10px;
    font-weight:700;
}}
.stButton>button:hover{{ border-color:{CYAN}; color:{CYAN}; }}
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


SHEET_MAP = {
    "PRS_Bulanan": "prs_monthly",
    "PRS_KPI": "prs_kpi",
    "PRS_Portofolio": "prs_portfolio",
    "PRS_Regional": "prs_regional",
    "AM_Performance": "am",
    "Action_Plan": "action",
}


SAVED_DATA_PATH = "saved_dashboard_data.xlsx"  # disimpan di server, dipakai bersama semua device

# Taruh file logo di folder "assets/" pada repo GitHub kamu, dengan nama-nama ini
LOGO_DIR = "assets"
LOGOS = [
    "logo_prs.png",
    "logo_divisi.png",
    "logo_telkom.png",
    "logo_danantara.png",
]


def render_logos():
    existing = [f for f in LOGOS if os.path.exists(os.path.join(LOGO_DIR, f))]
    if not existing:
        return
    cols = st.columns([2] + [1] * len(existing) + [2])
    for col, fname in zip(cols[1:-1], existing):
        with col:
            st.image(os.path.join(LOGO_DIR, fname), width=40)
    st.write("")


def _read_sheets_into(data, is_real, file_like, source_label):
    try:
        sheets = pd.read_excel(file_like, sheet_name=None)
        for sheet_name, key in SHEET_MAP.items():
            if sheet_name in sheets and not sheets[sheet_name].empty:
                data[key] = sheets[sheet_name]
                is_real[key] = True
        return True
    except Exception as e:
        st.sidebar.error(f"Gagal membaca {source_label}: {e}")
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
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        d["prs_monthly"].to_excel(writer, sheet_name="PRS_Bulanan", index=False)
        d["prs_kpi"].to_excel(writer, sheet_name="PRS_KPI", index=False)
        d["prs_portfolio"].to_excel(writer, sheet_name="PRS_Portofolio", index=False)
        d["prs_regional"].to_excel(writer, sheet_name="PRS_Regional", index=False)
        d["am"].to_excel(writer, sheet_name="AM_Performance", index=False)
        d["action"].to_excel(writer, sheet_name="Action_Plan", index=False)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# DATA PANEL — disembunyikan di balik ikon kecil (bukan sidebar), supaya tidak
# semua orang yang buka dashboard langsung lihat opsi upload/reset data.
# ----------------------------------------------------------------------------
_spacer, _icon_col = st.columns([20, 1])
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
            """)

        data, is_real, source = load_data(uploaded)
        n_real = sum(is_real.values())
        if source == "fresh_upload":
            st.success(f"✔ {n_real}/6 sheet tersimpan & aktif untuk SEMUA device yang buka dashboard ini.")
        elif source == "saved_on_server":
            st.info(f"📌 Memakai data tersimpan dari upload sebelumnya ({n_real}/6 sheet real).")
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
    st.markdown("<div style='text-align:center;padding-top:10px;'>", unsafe_allow_html=True)
    st.markdown("<div class='pill-title'>DASHBOARD OVERVIEW</div>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>A monthly performance overview to guide strategy and highlight progress</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
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
    st.markdown("## PRS <span style='color:#3fd6f0'>Performance</span>", unsafe_allow_html=True)

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
def render_am():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown("## AM <span style='color:#3fd6f0'>Performance</span>", unsafe_allow_html=True)

    am = data["am"].sort_values("Score", ascending=False).reset_index(drop=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg. Score AM", f"{am['Score'].mean():.1f}")
    c2.metric("Target Tercapai", f"{(am['Score'] >= am['Target']).sum()}/{len(am)}")
    c3.metric("Skor Tertinggi", f"{am['Score'].max():.1f}", am.iloc[0]["Name"])
    c4.metric("Skor Terendah", f"{am['Score'].min():.1f}", am.iloc[-1]["Name"])

    c1, c2 = st.columns([1.3, 1])
    with c1:
        st.markdown("#### Skor Performa per Account Manager")
        fig = go.Figure(go.Bar(x=am["Score"], y=am["Name"], orientation="h", marker_color=CYAN))
        fig.update_layout(**PLOTLY_LAYOUT, xaxis_range=[0, 100], height=320)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("#### Kompetensi — Top 3 AM")
        aspects = ["Komunikasi", "Negosiasi", "Kepatuhan", "KepuasanKlien", "Pelaporan"]
        fig = go.Figure()
        colors = [GOLD, CYAN, GREEN]
        for i in range(min(3, len(am))):
            row = am.iloc[i]
            fig.add_trace(go.Scatterpolar(
                r=[row[a] for a in aspects] + [row[aspects[0]]],
                theta=aspects + [aspects[0]],
                fill="toself", name=row["Name"], line_color=colors[i % 3],
            ))
        fig.update_layout(**PLOTLY_LAYOUT, polar=dict(
            bgcolor="rgba(0,0,0,0)",
            radialaxis=dict(visible=True, showticklabels=False, gridcolor="rgba(255,255,255,0.1)"),
            angularaxis=dict(gridcolor="rgba(255,255,255,0.1)"),
        ), height=320)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Leaderboard Account Manager")
    show = am.copy()
    show.insert(0, "Rank", range(1, len(show) + 1))
    show["Delta vs Target"] = (show["Score"] - show["Target"]).round(1)
    st.dataframe(
        show[["Rank", "Name", "Region", "Target", "Score", "Delta vs Target"]],
        use_container_width=True, hide_index=True,
    )


# ----------------------------------------------------------------------------
# ACTION PLAN AM PAGE
# ----------------------------------------------------------------------------
def render_action():
    st.button("← Overview", on_click=goto, args=("overview",))
    st.markdown("## Action Plan <span style='color:#3fd6f0'>AM</span>", unsafe_allow_html=True)

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

    styled = view.style.applymap(color_priority, subset=["Priority"]).applymap(color_status, subset=["Status"])
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

"""Aplikasi Streamlit perhitungan keekonomian migas."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from core.calculator import MigasInput, calculate_cashflow
from utils.excel_exporter import dataframe_to_xlsx_bytes


st.set_page_config(
    page_title="Aplikasi Perminyakan",
    page_icon="🛢️",
    layout="wide",
)

# ── Palet warna biru-pink pastel ──────────────────────────────────────────────
# Biru pastel  : #A8C8E8, #7BAFD4, #4A90C4
# Pink pastel  : #F4B8C8, #E8899A, #D4607A
# Background   : #F7F0F5 (lavender muda)
# Card bg      : #FFFFFF
# Sidebar      : #3D2B4F → #6B4C8A (ungu-biru gelap, serasi biru-pink)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background-color: #F7F0F5;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #3D2B4F 0%, #6B4C8A 100%);
    }
    [data-testid="stSidebar"] * {
        color: #F0E6F6 !important;
    }
    [data-testid="stSidebar"] .stNumberInput input,
    [data-testid="stSidebar"] .stSelectbox div {
        background-color: #5A3D7A !important;
        color: #FFFFFF !important;
        border: 1px solid #A882C8 !important;
        border-radius: 6px !important;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #F4B8C8 !important;
        border-bottom: 1px solid #A882C8;
        padding-bottom: 4px;
    }

    /* Hero banner */
    .hero-banner {
        background: linear-gradient(135deg, #4A90C4 0%, #7BAFD4 50%, #C490B8 100%);
        border-radius: 16px;
        padding: 36px 40px;
        margin-bottom: 28px;
        box-shadow: 0 4px 20px rgba(106,76,138,0.18);
    }
    .hero-banner h1 {
        color: #FFFFFF;
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 8px 0;
    }
    .hero-banner p {
        color: #EAF4FB;
        font-size: 1rem;
        margin: 0;
    }
    .hero-badge {
        display: inline-block;
        background: #F4B8C8;
        color: #3D2B4F;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 20px;
        margin-bottom: 12px;
        letter-spacing: 0.05em;
    }

    /* Metric cards */
    .metric-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 2px 10px rgba(106,76,138,0.10);
        border-left: 5px solid #7BAFD4;
        margin-bottom: 8px;
    }
    .metric-card.pink  { border-left-color: #E8899A; }
    .metric-card.blue  { border-left-color: #4A90C4; }
    .metric-card.lilac { border-left-color: #B090D0; }
    .metric-label {
        font-size: 0.78rem;
        font-weight: 600;
        color: #8B7AA0;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #3D2B4F;
    }

    /* Section heading */
    .section-heading {
        font-size: 1.05rem;
        font-weight: 700;
        color: #4A90C4;
        border-left: 4px solid #E8899A;
        padding-left: 10px;
        margin: 24px 0 12px 0;
    }

    /* Dataframe container */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        box-shadow: 0 2px 12px rgba(106,76,138,0.10);
    }

    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(90deg, #7BAFD4, #E8899A) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        transition: opacity 0.2s;
    }
    .stDownloadButton > button:hover {
        opacity: 0.85 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #EDE0F0;
        border-radius: 8px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 6px;
        color: #6B4C8A !important;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFFFFF !important;
        color: #4A90C4 !important;
        font-weight: 700 !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        background-color: #EDE0F0 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        color: #4A90C4 !important;
    }
    [data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border-radius: 8px !important;
    }
    [data-testid="stExpander"] p,
    [data-testid="stExpander"] td,
    [data-testid="stExpander"] th,
    [data-testid="stExpander"] li {
        color: #3D2B4F !important;
    }
    [data-testid="stExpander"] th {
        background-color: #EDE0F0 !important;
        font-weight: 700 !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Parameter Proyek")
    st.markdown("---")

    st.markdown("**Umum**")
    project_life = st.number_input("Jangka waktu proyek (tahun)", min_value=1, max_value=50, value=10)
    capital_investment = st.number_input("Investasi Capital ($M)", min_value=0.0, value=13000.0, step=100.0)
    non_capital_investment = st.number_input("Investasi Non Capital ($M)", min_value=0.0, value=8000.0, step=100.0)

    st.markdown("---")
    st.markdown("**Produksi Tahun 1–7**")
    default_production = [175.0, 201.0, 217.0, 198.0, 192.06, 186.29, 180.70]
    production_values: list[float] = []
    for index in range(7):
        value = st.number_input(
            f"Tahun ke-{index + 1} (Mbbl)",
            min_value=0.0,
            value=default_production[index],
            step=1.0,
        )
        production_values.append(value)

    st.markdown("---")
    st.markdown("**Asumsi Operasional**")
    production_decline_rate = st.number_input("Decline produksi/tahun (%)", min_value=0.0, max_value=100.0, value=3.0)
    oil_price = st.number_input("Harga minyak ($/bbl)", min_value=0.0, value=32.0, step=1.0)
    initial_opex = st.number_input("Opex awal ($M/tahun)", min_value=0.0, value=180.0, step=10.0)
    opex_escalation_rate = st.number_input("Kenaikan Opex/tahun (%)", min_value=0.0, value=2.5, step=0.1)
    opex_escalation_start_year = st.number_input("Opex mulai naik tahun ke-", min_value=1, max_value=50, value=4)

    st.markdown("---")
    st.markdown("**Depresiasi & Pajak**")
    depreciation_method = st.selectbox(
        "Metode depresiasi",
        ["Straight Line", "Double Declining Balance", "Sum of Years Digits"],
    )
    depreciation_life = st.number_input("Umur depresiasi (tahun)", min_value=1, max_value=50, value=20)
    tax_rate = st.number_input("Pajak (%)", min_value=0.0, max_value=100.0, value=51.0, step=1.0)

    st.markdown("---")
    st.caption("© 2026 · Keekonomian Migas")

# ── Hitung ────────────────────────────────────────────────────────────────────
params = MigasInput(
    project_life=int(project_life),
    capital_investment=float(capital_investment),
    non_capital_investment=float(non_capital_investment),
    production_year_1_to_7=production_values,
    production_decline_rate=float(production_decline_rate),
    oil_price=float(oil_price),
    initial_opex=float(initial_opex),
    opex_escalation_rate=float(opex_escalation_rate),
    opex_escalation_start_year=int(opex_escalation_start_year),
    depreciation_method=depreciation_method,
    depreciation_life=int(depreciation_life),
    tax_rate=float(tax_rate),
)

result_df = calculate_cashflow(params)

operating_rows = result_df[(result_df["Tahun"] != "Total") & (result_df["Tahun"] != 0)]
total_ncf    = pd.to_numeric(operating_rows["NCF Undiscounted ($M)"], errors="coerce").sum()
total_income = pd.to_numeric(operating_rows["Income ($M)"], errors="coerce").sum()
total_tax    = pd.to_numeric(operating_rows["Tax ($M)"], errors="coerce").sum()
total_opex   = pd.to_numeric(operating_rows["Opex ($M)"], errors="coerce").sum()

# ── Hero Banner ───────────────────────────────────────────────────────────────
st.markdown(
    f"""
    <div class="hero-banner">
        <div class="hero-badge">MIGAS ECONOMICS</div>
        <h1>Aplikasi Perminyakan</h1>
        <p>Perhitungan Keekonomian Lapangan Migas &nbsp;·&nbsp; Proyeksi {int(project_life)} tahun &nbsp;·&nbsp; Harga minyak ${float(oil_price):,.0f}/bbl
        &nbsp;·&nbsp; Metode depresiasi: {depreciation_method}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Metric Cards ──────────────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(
        f"""<div class="metric-card blue">
            <div class="metric-label">Total NCF</div>
            <div class="metric-value">${total_ncf:,.1f}M</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f"""<div class="metric-card">
            <div class="metric-label">Total Income</div>
            <div class="metric-value">${total_income:,.1f}M</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f"""<div class="metric-card pink">
            <div class="metric-label">Total Pajak</div>
            <div class="metric-value">${total_tax:,.1f}M</div>
        </div>""",
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        f"""<div class="metric-card lilac">
            <div class="metric-label">Total Opex</div>
            <div class="metric-value">${total_opex:,.1f}M</div>
        </div>""",
        unsafe_allow_html=True,
    )

# ── Chart helpers ─────────────────────────────────────────────────────────────
chart_df = operating_rows.copy()
chart_df["Tahun"] = pd.to_numeric(chart_df["Tahun"], errors="coerce")
for col in ["Income ($M)", "Opex ($M)", "Tax ($M)", "NCF Undiscounted ($M)"]:
    chart_df[col] = pd.to_numeric(chart_df[col], errors="coerce")

PLOT_BG   = "#FDF6FB"   # pink lavender sangat muda
PAPER_BG  = "#FDF6FB"
GRID_COL  = "#E8D8F0"
FONT_COL  = "#3D2B4F"

def base_layout(title: str) -> dict:
    return dict(
        title=dict(text=title, font=dict(color=FONT_COL, size=14)),
        plot_bgcolor=PLOT_BG,
        paper_bgcolor=PAPER_BG,
        font=dict(family="Inter, sans-serif", color=FONT_COL),
        xaxis=dict(
            title="Tahun",
            gridcolor=GRID_COL,
            linecolor=GRID_COL,
            tickfont=dict(color=FONT_COL),
        ),
        yaxis=dict(
            title="$M",
            gridcolor=GRID_COL,
            linecolor=GRID_COL,
            tickfont=dict(color=FONT_COL),
        ),
        legend=dict(
            bgcolor="rgba(255,255,255,0.7)",
            bordercolor=GRID_COL,
            borderwidth=1,
        ),
        margin=dict(l=50, r=20, t=50, b=40),
        bargap=0.25,
    )

# ── Charts ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-heading">Grafik per Tahun</div>', unsafe_allow_html=True)

tab_chart1, tab_chart2 = st.tabs(["Income vs Opex vs Tax", "NCF per Tahun"])

with tab_chart1:
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(
        x=chart_df["Tahun"], y=chart_df["Income ($M)"],
        name="Income", marker_color="#7BAFD4",
    ))
    fig1.add_trace(go.Bar(
        x=chart_df["Tahun"], y=chart_df["Opex ($M)"],
        name="Opex", marker_color="#B090D0",
    ))
    fig1.add_trace(go.Bar(
        x=chart_df["Tahun"], y=chart_df["Tax ($M)"],
        name="Tax", marker_color="#F4B8C8",
    ))
    fig1.update_layout(**base_layout("Income, Opex, dan Tax per Tahun"), barmode="group")
    st.plotly_chart(fig1, use_container_width=True)

with tab_chart2:
    ncf_vals = chart_df["NCF Undiscounted ($M)"]
    bar_colors = ["#7BAFD4" if v >= 0 else "#E8899A" for v in ncf_vals]
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(
        x=chart_df["Tahun"], y=ncf_vals,
        name="NCF", marker_color=bar_colors,
    ))
    fig2.add_hline(y=0, line_color="#B090D0", line_width=1.2, line_dash="dot")
    fig2.update_layout(**base_layout("NCF Undiscounted per Tahun"))
    st.plotly_chart(fig2, use_container_width=True)

# ── Tabel ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-heading">Tabel Cashflow</div>', unsafe_allow_html=True)

numeric_cols = [
    "Produksi (Mbbl)", "Income ($M)", "Capital ($M)", "Non Capital ($M)",
    "Opex ($M)", "Di ($M)", "Taxable Income ($M)", "Tax ($M)", "NCF Undiscounted ($M)",
]

st.dataframe(
    result_df.style.format(
        {col: "{:,.2f}" for col in numeric_cols},
        na_rep="-",
    ).apply(
        lambda row: [
            "background-color: #F9E4EE; color: #7D1040; font-weight: 700;" if row["Tahun"] == "Total"
            else ""
            for _ in row
        ],
        axis=1,
    ),
    use_container_width=True,
    hide_index=True,
    height=(len(result_df) + 1) * 35 + 10,
)

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-heading">Export Data</div>', unsafe_allow_html=True)

xlsx_data = dataframe_to_xlsx_bytes(result_df)
st.download_button(
    label="Download Hasil (.xlsx)",
    data=xlsx_data,
    file_name="hasil_perhitungan_migas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=False,
)

# ── Rumus ─────────────────────────────────────────────────────────────────────
with st.expander("Rumus yang digunakan"):
    st.markdown(
        """
        | Variabel | Rumus |
        |---|---|
        | **Income** | Produksi × Harga Minyak |
        | **Taxable Income** | Income − Opex − Depresiasi (Dᵢ) |
        | **Tax** | Taxable Income × Tarif Pajak *(jika positif)* |
        | **NCF Undiscounted** | Income − Opex − Tax − Dᵢ |
        | **NCF₀** | −(Capital + Non Capital) |
        """
    )

"""Aplikasi Streamlit perhitungan keekonomian migas."""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from core.calculator import MigasInput, calculate_cashflow
from utils.excel_exporter import dataframe_to_xlsx_bytes


st.set_page_config(
    page_title="Perhitungan Keekonomian Migas",
    page_icon="🛢️",
    layout="wide",
)

st.title("Perhitungan Keekonomian Lapangan Migas")
st.write(
    "Masukkan data proyek, lalu hasil perhitungan akan tampil langsung sebagai tabel "
    "dan bisa di-download dalam format Excel."
)

with st.sidebar:
    st.header("Input Utama")

    project_life = st.number_input("Jangka waktu proyek (tahun)", min_value=1, max_value=50, value=10)
    capital_investment = st.number_input("Investasi Capital ($M)", min_value=0.0, value=13000.0, step=100.0)
    non_capital_investment = st.number_input("Investasi Non Capital ($M)", min_value=0.0, value=8000.0, step=100.0)

    st.subheader("Produksi Tahun 1 s.d. 7")
    default_production = [175.0, 201.0, 217.0, 198.0, 192.06, 186.29, 180.70]
    production_values: list[float] = []
    for index in range(7):
        value = st.number_input(
            f"Produksi tahun ke-{index + 1} (Mbbl)",
            min_value=0.0,
            value=default_production[index],
            step=1.0,
        )
        production_values.append(value)

    st.subheader("Asumsi Operasional")
    production_decline_rate = st.number_input("Decline produksi per tahun (%)", min_value=0.0, max_value=100.0, value=3.0)
    oil_price = st.number_input("Harga rata-rata minyak ($/bbl)", min_value=0.0, value=32.0, step=1.0)
    initial_opex = st.number_input("Opex awal ($M/tahun)", min_value=0.0, value=180.0, step=10.0)
    opex_escalation_rate = st.number_input("Kenaikan Opex per tahun (%)", min_value=0.0, value=2.5, step=0.1)
    opex_escalation_start_year = st.number_input("Opex mulai naik pada tahun ke-", min_value=1, max_value=50, value=4)

    st.subheader("Depresiasi dan Pajak")
    depreciation_method = st.selectbox(
        "Metode depresiasi",
        ["Straight Line", "Double Declining Balance", "Sum of Years Digits"],
    )
    depreciation_life = st.number_input("Umur depresiasi (tahun)", min_value=1, max_value=50, value=20)
    tax_rate = st.number_input("Pajak (%)", min_value=0.0, max_value=100.0, value=51.0, step=1.0)

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

st.subheader("Hasil Perhitungan")
st.dataframe(
    result_df.style.format(
        {
            "Produksi (Mbbl)": "{:,.2f}",
            "Income ($M)": "{:,.2f}",
            "Capital ($M)": "{:,.2f}",
            "Non Capital ($M)": "{:,.2f}",
            "Opex ($M)": "{:,.2f}",
            "Di ($M)": "{:,.2f}",
            "Taxable Income ($M)": "{:,.2f}",
            "Tax ($M)": "{:,.2f}",
            "NCF Undiscounted ($M)": "{:,.2f}",
        }
    ),
    use_container_width=True,
    hide_index=True,
)

col1, col2, col3 = st.columns(3)
with col1:
    operating_rows = result_df[result_df["Tahun"] != "Total"]
    total_ncf = pd.to_numeric(operating_rows["NCF Undiscounted ($M)"], errors="coerce").sum()
    st.metric("Total NCF ($M)", f"{total_ncf:,.2f}")
with col2:
    total_income = pd.to_numeric(operating_rows["Income ($M)"], errors="coerce").sum()
    st.metric("Total Income ($M)", f"{total_income:,.2f}")
with col3:
    total_tax = pd.to_numeric(operating_rows["Tax ($M)"], errors="coerce").sum()
    st.metric("Total Tax ($M)", f"{total_tax:,.2f}")

xlsx_data = dataframe_to_xlsx_bytes(result_df)
st.download_button(
    label="Download Hasil XLSX",
    data=xlsx_data,
    file_name="hasil_perhitungan_migas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("Rumus yang digunakan"):
    st.markdown(
        """
        - **Income** = Produksi x Harga Minyak
        - **Taxable Income** = Income - Opex - Depresiasi
        - **Tax** = Taxable Income x Pajak, jika Taxable Income positif
        - **NCF Undiscounted** = Taxable Income - Tax
        - Pada tahun ke-0, NCF = -(Capital + Non Capital)
        """
    )

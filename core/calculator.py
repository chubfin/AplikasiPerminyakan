"""Modul utama perhitungan keekonomian migas."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from core.depreciation import calculate_depreciation


@dataclass
class MigasInput:
    project_life: int
    capital_investment: float
    non_capital_investment: float
    production_year_1_to_7: list[float]
    production_decline_rate: float
    oil_price: float
    initial_opex: float
    opex_escalation_rate: float
    opex_escalation_start_year: int
    depreciation_method: str
    depreciation_life: int
    tax_rate: float

    @property
    def total_investment(self) -> float:
        return self.capital_investment + self.non_capital_investment


def _production_for_year(params: MigasInput, year: int, previous_production: float | None) -> float:
    """Mengambil produksi input atau menghitung decline otomatis."""
    if year <= len(params.production_year_1_to_7):
        return float(params.production_year_1_to_7[year - 1])

    if previous_production is None:
        return 0.0

    return previous_production * (1 - params.production_decline_rate / 100)


def _opex_for_year(params: MigasInput, year: int, previous_opex: float | None) -> float:
    """Menghitung Opex tahunan."""
    if year <= 0:
        return 0.0

    if year < params.opex_escalation_start_year:
        return params.initial_opex

    if year == params.opex_escalation_start_year:
        if params.opex_escalation_start_year <= 1:
            return params.initial_opex
        return params.initial_opex * (1 + params.opex_escalation_rate / 100)

    base = previous_opex if previous_opex is not None else params.initial_opex
    return base * (1 + params.opex_escalation_rate / 100)


def calculate_cashflow(params: MigasInput) -> pd.DataFrame:
    """Menghasilkan tabel cashflow keekonomian migas."""
    rows: list[dict[str, float | int | str]] = []

    rows.append(
        {
            "Tahun": 0,
            "Produksi (Mbbl)": 0.0,
            "Income ($M)": 0.0,
            "Capital ($M)": params.capital_investment,
            "Non Capital ($M)": params.non_capital_investment,
            "Opex ($M)": 0.0,
            "Di ($M)": 0.0,
            "Taxable Income ($M)": 0.0,
            "Tax ($M)": 0.0,
            "NCF Undiscounted ($M)": -params.total_investment,
        }
    )

    previous_production: float | None = None
    previous_opex: float | None = None

    for year in range(1, params.project_life + 1):
        production = _production_for_year(params, year, previous_production)
        income = production * params.oil_price
        opex = _opex_for_year(params, year, previous_opex)
        depreciation = calculate_depreciation(
            method=params.depreciation_method,
            asset_value=params.total_investment,
            project_year=year,
            useful_life=params.depreciation_life,
        )
        taxable_income = income - opex - depreciation
        tax = max(taxable_income, 0) * params.tax_rate / 100
        ncf = taxable_income - tax

        rows.append(
            {
                "Tahun": year,
                "Produksi (Mbbl)": production,
                "Income ($M)": income,
                "Capital ($M)": 0.0,
                "Non Capital ($M)": 0.0,
                "Opex ($M)": opex,
                "Di ($M)": depreciation,
                "Taxable Income ($M)": taxable_income,
                "Tax ($M)": tax,
                "NCF Undiscounted ($M)": ncf,
            }
        )

        previous_production = production
        previous_opex = opex

    total_row = {column: "" for column in rows[0].keys()}
    total_row["Tahun"] = "Total"
    for column in [
        "Produksi (Mbbl)",
        "Income ($M)",
        "Capital ($M)",
        "Non Capital ($M)",
        "Opex ($M)",
        "Di ($M)",
        "Taxable Income ($M)",
        "Tax ($M)",
        "NCF Undiscounted ($M)",
    ]:
        total_row[column] = sum(float(row[column]) for row in rows)

    rows.append(total_row)
    return pd.DataFrame(rows)

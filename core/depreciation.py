"""Modul perhitungan depresiasi."""

from __future__ import annotations


def calculate_depreciation(
    method: str,
    asset_value: float,
    project_year: int,
    useful_life: int,
) -> float:
    """Menghitung depresiasi tahunan.

    Args:
        method: Metode depresiasi.
        asset_value: Nilai aset yang didepresiasi.
        project_year: Tahun proyek, dimulai dari 1.
        useful_life: Umur manfaat aset.

    Returns:
        Nilai depresiasi untuk tahun terkait.
    """
    if asset_value <= 0 or useful_life <= 0 or project_year <= 0:
        return 0.0

    method = method.lower().strip()

    if project_year > useful_life:
        return 0.0

    if method == "straight line":
        return asset_value / useful_life

    if method == "double declining balance":
        rate = 2 / useful_life
        book_value_beginning = asset_value * ((1 - rate) ** (project_year - 1))
        depreciation = book_value_beginning * rate
        remaining_years = useful_life - project_year + 1
        straight_line_remaining = book_value_beginning / remaining_years
        return min(max(depreciation, straight_line_remaining), book_value_beginning)

    if method == "sum of years digits":
        denominator = useful_life * (useful_life + 1) / 2
        remaining_life = useful_life - project_year + 1
        return asset_value * remaining_life / denominator

    raise ValueError(f"Metode depresiasi tidak dikenali: {method}")

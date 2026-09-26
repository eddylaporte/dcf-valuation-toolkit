"""
Nettoyage des états financiers bruts (yfinance) : uniformisation des
colonnes, extraction robuste de lignes par libellés multiples, calcul de
la marge d'EBIT historique.
"""

import logging

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)


def clean_statement(df: pd.DataFrame) -> pd.DataFrame:
    """Uniformise un état financier : colonnes triées chronologiquement, NaN gérés."""
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.copy()
    df.columns = [c.strftime("%Y") if isinstance(c, pd.Timestamp) else str(c) for c in df.columns]
    df = df.sort_index(axis=1)
    df = df.apply(pd.to_numeric, errors="coerce")
    return df.dropna(how="all")


def build_clean_dataset(raw: dict) -> dict:
    return {
        "income_stmt": clean_statement(raw["income_stmt"]),
        "balance_sheet": clean_statement(raw["balance_sheet"]),
        "cashflow": clean_statement(raw["cashflow"]),
    }


def get_row(df: pd.DataFrame, candidates: list[str]) -> pd.Series | None:
    """Retourne la première ligne trouvée parmi plusieurs libellés possibles."""
    for name in candidates:
        if name in df.index:
            return df.loc[name]
    return None


def get_latest_value(df: pd.DataFrame, candidates: list[str], default=np.nan):
    row = get_row(df, candidates)
    if row is not None:
        series = row.dropna()
        if not series.empty:
            return series.iloc[-1]
    return default


def get_ebit_margin_trend(income_stmt: pd.DataFrame, lookback_years: int = 3) -> list[tuple[str, float]]:
    """
    Retourne la marge d'EBIT année par année (label d'exercice, marge) sur les
    derniers exercices disponibles — pour documenter explicitement si la
    moyenne utilisée dans le DCF masque une tendance haussière ou baissière,
    plutôt que de ne présenter qu'un chiffre unique.
    """
    ebit_row = get_row(income_stmt, ["Operating Income", "OperatingIncome", "EBIT"])
    revenue_row = get_row(income_stmt, ["Total Revenue", "TotalRevenue"])
    if ebit_row is None or revenue_row is None:
        return []
    merged = pd.DataFrame({"ebit": ebit_row, "revenue": revenue_row}).dropna()
    if merged.empty:
        return []
    merged = merged.tail(lookback_years)
    return [(str(year), float(row["ebit"] / row["revenue"])) for year, row in merged.iterrows()]


def compute_historical_ebit_margin(income_stmt: pd.DataFrame, lookback_years: int = 3,
                                    weighted: bool = True) -> float | None:
    """
    Calcule la marge d'EBIT retenue sur les derniers exercices disponibles, à
    partir des données réelles — jamais une hypothèse arbitraire.

    Par défaut (weighted=True), la moyenne est pondérée en donnant plus de poids
    aux exercices récents (poids croissants 1, 2, 3...) plutôt qu'une moyenne
    simple : si la marge suit une tendance marquée, une moyenne simple la sous-
    ou sur-estime systématiquement par rapport à la trajectoire réelle de
    l'entreprise. La pondération corrige ce biais au lieu de se contenter de le
    signaler. Passer weighted=False pour obtenir la moyenne simple (utile pour
    comparaison ou audit).

    Retourne None si les lignes nécessaires sont absentes.
    """
    trend = get_ebit_margin_trend(income_stmt, lookback_years)
    if not trend:
        return None
    margins = [m for _, m in trend]
    if weighted and len(margins) > 1:
        weights = list(range(1, len(margins) + 1))
        return float(sum(m * w for m, w in zip(margins, weights)) / sum(weights))
    return float(sum(margins) / len(margins))
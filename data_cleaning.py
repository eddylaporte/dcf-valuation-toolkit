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


def compute_historical_ebit_margin(income_stmt: pd.DataFrame, lookback_years: int = 3) -> float | None:
    """
    Calcule la marge d'EBIT moyenne sur les derniers exercices disponibles,
    à partir des données réelles, plutôt que d'utiliser une hypothèse arbitraire.
    Retourne None si les lignes nécessaires sont absentes.
    """
    ebit_row = get_row(income_stmt, ["Operating Income", "OperatingIncome", "EBIT"])
    revenue_row = get_row(income_stmt, ["Total Revenue", "TotalRevenue"])
    if ebit_row is None or revenue_row is None:
        return None

    merged = pd.DataFrame({"ebit": ebit_row, "revenue": revenue_row}).dropna()
    if merged.empty:
        return None

    merged = merged.tail(lookback_years)
    margins = merged["ebit"] / merged["revenue"]
    logger.info("Marges d'EBIT historiques (%d derniers exercices) : %s",
                len(margins), ", ".join(f"{m:.1%}" for m in margins))
    return float(margins.mean())
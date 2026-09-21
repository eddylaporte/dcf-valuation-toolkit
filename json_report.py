"""
Export des résultats clés d'un run vers un fichier JSON — pensé pour être
réutilisé ailleurs (page web, autre outil) sans reparser la sortie console.
"""

import json
import logging

import pandas as pd

from valuation_config import ValuationAssumptions

logger = logging.getLogger(__name__)


def export_results_json(path: str, ticker: str, peers: list[str], current_price: float | None,
                         dcf_result: dict, scenario_values: dict, comps_df: pd.DataFrame,
                         comps_avg: float | None, consensus: dict, recommendation: dict,
                         wacc: float, ebit_margin_base: float, assumptions: ValuationAssumptions,
                         ff_data: dict) -> None:
    """
    Exporte les chiffres clés du run dans un fichier JSON (un seul fichier texte,
    lisible tel quel) — pensé pour être réinjecté ensuite dans une page web ou tout
    autre outil, sans avoir à reparser la sortie console.
    """
    data = {
        "ticker": ticker,
        "peers": peers,
        "current_price": current_price,
        "dcf": {
            "value_per_share_base": dcf_result["value_per_share"],
            "enterprise_value_m": dcf_result["enterprise_value"] / 1e6,
            "equity_value_m": dcf_result["equity_value"] / 1e6,
            "wacc": wacc,
            "ebit_margin_base": ebit_margin_base,
            "terminal_growth_base": assumptions.terminal_growth_base,
        },
        "scenarios": {
            "bear": scenario_values.get(1),
            "base": scenario_values.get(2),
            "bull": scenario_values.get(3),
        },
        "comps": {
            "average_value_per_share": comps_avg,
            "methods": comps_df.to_dict(orient="records"),
        },
        "consensus": consensus,
        "recommendation": recommendation,
        "football_field": ff_data,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)
    logger.info("Résultats clés exportés en JSON : %s", path)
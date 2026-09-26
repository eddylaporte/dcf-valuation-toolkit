"""
Configuration de la valorisation : hypothèses par défaut (WACC, croissance,
marges par scénario) et sources documentées pour chaque hypothèse macro.
"""

from dataclasses import dataclass, field


@dataclass
class ValuationAssumptions:
    risk_free_rate: float = 0.042
    market_risk_premium: float = 0.05
    cost_of_debt: float = 0.035
    tax_rate: float = 0.25
    projection_years: int = 5

    # Croissance du CA par scénario (une valeur par année de projection)
    revenue_growth_bear: list = field(default_factory=lambda: [0.00, -0.005, -0.01, -0.01, -0.005])
    revenue_growth_base: list = field(default_factory=lambda: [0.02, 0.02, 0.015, 0.015, 0.01])
    revenue_growth_bull: list = field(default_factory=lambda: [0.04, 0.035, 0.03, 0.025, 0.02])

    # La marge d'EBIT "Base" est calculée depuis l'historique (voir
    # compute_historical_ebit_margin) ; Bear/Bull appliquent un delta.
    ebit_margin_delta_bear: float = -0.02
    ebit_margin_delta_bull: float = 0.02
    ebit_margin_override: float = None  # si renseigné, force la marge "Base"

    terminal_growth_bear: float = 0.01
    terminal_growth_base: float = 0.02
    terminal_growth_bull: float = 0.03

    default_scenario: int = 2  # 1 = Bear, 2 = Base, 3 = Bull

    capex_pct_revenue: float = 0.08
    da_pct_revenue: float = 0.06
    nwc_pct_revenue_change: float = 0.005

    beta_fallback: float = 0.95
    beta_plausible_range: tuple = (0.3, 2.5)
    beta_peer_divergence_threshold: float = 0.35  # écart relatif au-delà duquel le beta
                                                    # brut est jugé suspect même s'il est
                                                    # dans la plage plausible générique, et
                                                    # remplacé par le beta moyen des pairs


# Sources par défaut des hypothèses macro/marché — à vérifier et actualiser
# à la date de l'analyse. Reprises dans la note Word et dans les notes de
# cellules Excel pour que le lecteur sache d'où vient chaque chiffre.
ASSUMPTION_SOURCES = {
    "Taux sans risque": "10Y US Treasury — U.S. Department of the Treasury "
                         "(home.treasury.gov/resource-center/data-chart-center/interest-rates) "
                         "ou FRED, série DGS10.",
    "Prime de risque marché": "Equity Risk Premium (US) — Damodaran, NYU Stern "
                               "(pages.stern.nyu.edu/~adamodar).",
    "Beta": "yfinance (5Y monthly vs S&P 500). À recouper avec Bloomberg/Refinitiv si la "
            "valeur semble aberrante, ou avec le beta moyen des comparables.",
    "Taux d'imposition": "Taux effectif observé sur les derniers comptes annuels de la "
                          "société (10-K / 20-F / document d'enregistrement universel).",
    "Coût de la dette": "Approximé par le taux d'intérêt moyen sur la dette existante "
                         "(notes annexes des états financiers) ou le spread de crédit observé.",
}
"""
Cœur de calcul de la valorisation : WACC (CAPM), projection des FCF et DCF
par scénario, valorisation par comparables, football field et règle de
recommandation. Calculs Python purs, indépendants du classeur Excel (qui
recalcule tout via ses propres formules).
"""

import numpy as np
import pandas as pd

from valuation_config import ValuationAssumptions


def compute_wacc(beta: float, equity_value: float, debt_value: float,
                  assumptions: ValuationAssumptions) -> tuple[float, float]:
    cost_of_equity = assumptions.risk_free_rate + beta * assumptions.market_risk_premium
    total = equity_value + debt_value
    if total == 0:
        return cost_of_equity, cost_of_equity
    weight_equity = equity_value / total
    weight_debt = debt_value / total
    wacc = (weight_equity * cost_of_equity
            + weight_debt * assumptions.cost_of_debt * (1 - assumptions.tax_rate))
    return wacc, cost_of_equity


def project_fcf(base_revenue: float, growth_path: list[float], ebit_margin: float,
                 tax_rate: float, capex_pct_revenue: float, da_pct_revenue: float,
                 nwc_pct_revenue_change: float) -> pd.DataFrame:
    """Construit les flux de trésorerie disponibles (FCF) projetés pour un scénario donné."""
    years = [f"N+{i + 1}" for i in range(len(growth_path))]
    revenues, revenue = [], base_revenue
    for growth in growth_path:
        revenue *= (1 + growth)
        revenues.append(revenue)

    ebit = [r * ebit_margin for r in revenues]
    nopat = [e * (1 - tax_rate) for e in ebit]
    da = [r * da_pct_revenue for r in revenues]
    capex = [r * capex_pct_revenue for r in revenues]
    delta_nwc = [r * nwc_pct_revenue_change for r in revenues]
    fcf = [n + d - c - w for n, d, c, w in zip(nopat, da, capex, delta_nwc)]

    return pd.DataFrame({
        "Année": years, "Revenus": revenues, "EBIT": ebit, "NOPAT": nopat,
        "D&A": da, "Capex": capex, "Delta_BFR": delta_nwc, "FCF": fcf,
    })


def run_dcf(fcf_df: pd.DataFrame, wacc: float, terminal_growth: float,
            net_debt: float, shares_outstanding: float) -> dict:
    """Actualise les FCF et la valeur terminale, puis dérive la valeur par action."""
    fcf = fcf_df["FCF"].values
    discount_factors = np.array([1 / (1 + wacc) ** (i + 1) for i in range(len(fcf))])
    pv_fcf = fcf * discount_factors

    terminal_value = fcf[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value * discount_factors[-1]

    enterprise_value = pv_fcf.sum() + pv_terminal
    equity_value = enterprise_value - net_debt
    value_per_share = equity_value / shares_outstanding if shares_outstanding else np.nan

    return {
        "pv_fcf_sum": pv_fcf.sum(),
        "terminal_value": terminal_value,
        "pv_terminal_value": pv_terminal,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "value_per_share": value_per_share,
    }


def scenario_inputs(assumptions: ValuationAssumptions, ebit_margin_base: float,
                     scenario: int) -> tuple[list[float], float, float]:
    """Retourne (croissance du CA, marge d'EBIT, croissance terminale) pour un scénario (1/2/3)."""
    growth_paths = {1: assumptions.revenue_growth_bear, 2: assumptions.revenue_growth_base,
                     3: assumptions.revenue_growth_bull}
    terminal_growths = {1: assumptions.terminal_growth_bear, 2: assumptions.terminal_growth_base,
                         3: assumptions.terminal_growth_bull}
    margin_deltas = {1: assumptions.ebit_margin_delta_bear, 2: 0.0, 3: assumptions.ebit_margin_delta_bull}
    if scenario not in growth_paths:
        raise ValueError("scenario doit être 1 (Bear), 2 (Base) ou 3 (Bull)")
    return growth_paths[scenario], ebit_margin_base + margin_deltas[scenario], terminal_growths[scenario]


def run_scenario_dcf(total_revenue: float, ebit_margin_base: float, net_debt: float,
                      shares_outstanding: float, wacc: float, assumptions: ValuationAssumptions,
                      scenario: int) -> dict:
    """Calcule le DCF complet pour un scénario donné (1=Bear, 2=Base, 3=Bull)."""
    growth, ebit_margin, terminal_growth = scenario_inputs(assumptions, ebit_margin_base, scenario)
    fcf_df = project_fcf(total_revenue, growth, ebit_margin, assumptions.tax_rate,
                          assumptions.capex_pct_revenue, assumptions.da_pct_revenue,
                          assumptions.nwc_pct_revenue_change)
    return run_dcf(fcf_df, wacc, terminal_growth, net_debt, shares_outstanding)


# ----------------------------------------------------------------------
# Comparables (calcul Python informatif — le classeur Excel recalcule via MEDIAN())
# ----------------------------------------------------------------------

def run_comps(target_metrics: dict, peers_df: pd.DataFrame) -> pd.DataFrame:
    """Applique les multiples médians des pairs aux métriques de la cible."""
    medians = peers_df[["trailing_pe", "ev_to_ebitda", "ev_to_revenue"]].median(numeric_only=True)
    return pd.DataFrame({
        "Méthode": ["EV/EBITDA", "EV/Sales", "P/E"],
        "Multiple médian peers": [medians["ev_to_ebitda"], medians["ev_to_revenue"], medians["trailing_pe"]],
        "Valeur implicite": [
            medians["ev_to_ebitda"] * target_metrics["ebitda"],
            medians["ev_to_revenue"] * target_metrics["revenue"],
            medians["trailing_pe"] * target_metrics["net_income"],
        ],
    })


def comps_per_share(comps_df: pd.DataFrame, net_debt: float, shares_outstanding: float) -> list[float]:
    """Convertit chaque valeur implicite des comps en valeur par action (EV -> equity pour EV/EBITDA et EV/Sales)."""
    values = []
    for _, row in comps_df.iterrows():
        if row["Méthode"] == "P/E":
            values.append(row["Valeur implicite"] / shares_outstanding)
        else:
            values.append((row["Valeur implicite"] - net_debt) / shares_outstanding)
    return values


def compute_football_field_data(fifty_two_week_low: float | None, fifty_two_week_high: float | None,
                                 scenario_values: list[float], comps_values: list[float]) -> dict:
    """Construit les fourchettes (low, high) par méthode pour le graphique football field."""
    data = {
        "DCF (Bear-Bull)": (min(scenario_values), max(scenario_values)),
        "Comparables (min-max)": (min(comps_values), max(comps_values)),
    }
    if fifty_two_week_low and fifty_two_week_high:
        data["52 semaines (marché)"] = (fifty_two_week_low, fifty_two_week_high)
    return data


DIVERGENCE_EXCLUSION_THRESHOLD = 0.30  # points d'écart vs le DCF au-delà desquels une
                                        # méthode est jugée trop peu fiable pour entrer
                                        # dans la moyenne (voir generate_recommendation)


def generate_recommendation(dcf_value: float, comps_avg: float, consensus_target: float | None,
                             current_price: float | None) -> dict:
    """
    Dérive une recommandation (Achat / Conserver / Vente) à partir de l'upside moyen
    implicite des méthodes de valorisation disponibles par rapport au prix de marché.

    Le DCF sert d'ancre (c'est la méthode la plus rigoureuse, fondée sur les flux de
    trésorerie réels de l'entreprise). Toute autre méthode (comparables, consensus) qui
    s'écarte de l'upside du DCF de plus de DIVERGENCE_EXCLUSION_THRESHOLD (30 points de
    pourcentage par défaut) est exclue de la moyenne plutôt que d'y être mélangée : une
    méthode connue pour être bruyante (un multiple de pair faussé par un résultat
    exceptionnel, par exemple) ne doit pas pouvoir, à elle seule, faire basculer la
    recommandation dans un sens contraire à ce que suggèrent les flux de trésorerie.

    Règle de décision sur la moyenne des méthodes retenues :
        upside moyen > +15 %  -> Achat
        upside moyen < -10 %  -> Vente
        entre les deux        -> Conserver
    C'est un point de départ mécanique et transparent, pas un jugement définitif : il ne
    remplace pas une analyse qualitative des catalyseurs et risques propres à la thèse
    d'investissement.
    """
    if not current_price:
        return {"label": "Conserver (Hold)", "avg_upside": None, "upsides": {}, "excluded": {}}

    dcf_upside = (dcf_value - current_price) / current_price
    upsides = {"DCF": dcf_upside}
    excluded = {}

    candidates = {}
    if comps_avg:
        candidates["Comparables"] = (comps_avg - current_price) / current_price
    if consensus_target:
        candidates["Consensus analystes"] = (consensus_target - current_price) / current_price

    for method, upside in candidates.items():
        if abs(upside - dcf_upside) > DIVERGENCE_EXCLUSION_THRESHOLD:
            excluded[method] = upside
        else:
            upsides[method] = upside

    avg_upside = sum(upsides.values()) / len(upsides)
    if avg_upside > 0.15:
        label = "Achat (Buy)"
    elif avg_upside < -0.10:
        label = "Vente (Sell)"
    else:
        label = "Conserver (Hold)"
    return {"label": label, "avg_upside": avg_upside, "upsides": upsides, "excluded": excluded}
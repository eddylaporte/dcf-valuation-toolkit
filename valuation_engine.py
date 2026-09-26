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


def compute_peer_average_beta(peers_df: pd.DataFrame) -> float | None:
    """Beta moyen des pairs, utilisé comme référence de contrôle indépendante du beta brut."""
    valid = peers_df["beta"].dropna() if "beta" in peers_df else pd.Series(dtype=float)
    if valid.empty:
        return None
    return float(valid.mean())


def resolve_beta(beta_raw: float | None, peer_avg_beta: float | None,
                  assumptions: ValuationAssumptions) -> tuple[float, str]:
    """
    Détermine le beta à retenir pour le WACC, avec une vraie vérification croisée
    plutôt qu'un simple test de plage plausible générique : un beta techniquement
    "dans la plage" (ex: 0.3-2.5) peut rester peu crédible pour une société donnée
    s'il s'écarte fortement du beta moyen de ses propres pairs. Retourne (beta
    retenu, explication) — l'explication est destinée à la note de synthèse pour
    que le choix soit traçable, pas une boîte noire.
    """
    lo, hi = assumptions.beta_plausible_range
    out_of_range = beta_raw is None or not (lo <= beta_raw <= hi)

    if out_of_range:
        if peer_avg_beta is not None:
            return peer_avg_beta, (
                f"Beta brut ({beta_raw}) hors plage plausible {assumptions.beta_plausible_range} "
                f"-> beta moyen des pairs retenu ({peer_avg_beta:.2f})."
            )
        return assumptions.beta_fallback, (
            f"Beta brut ({beta_raw}) hors plage plausible et beta moyen des pairs indisponible "
            f"-> valeur de repli sectorielle retenue ({assumptions.beta_fallback:.2f}), à vérifier "
            "manuellement (Bloomberg/Refinitiv) avant toute diffusion."
        )

    if peer_avg_beta is not None and peer_avg_beta > 0:
        relative_gap = abs(beta_raw - peer_avg_beta) / peer_avg_beta
        if relative_gap > assumptions.beta_peer_divergence_threshold:
            return peer_avg_beta, (
                f"Beta brut ({beta_raw:.2f}) dans la plage plausible mais s'écarte de {relative_gap:.0%} "
                f"du beta moyen des pairs ({peer_avg_beta:.2f}) — écart jugé trop important pour être "
                "retenu tel quel -> beta moyen des pairs retenu par prudence. À vérifier auprès d'une "
                "source indépendante (Bloomberg/Refinitiv) avant diffusion finale."
            )
        return beta_raw, (
            f"Beta brut ({beta_raw:.2f}) cohérent avec le beta moyen des pairs ({peer_avg_beta:.2f}, "
            f"écart de {relative_gap:.0%}) -> retenu tel quel."
        )

    return beta_raw, (f"Beta brut ({beta_raw:.2f}) dans la plage plausible ; beta moyen des pairs "
                       "indisponible pour contrôle croisé.")


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


def comps_per_share(comps_df: pd.DataFrame, net_debt: float, shares_outstanding: float) -> dict[str, float]:
    """Convertit chaque valeur implicite des comps en valeur par action (EV -> equity pour EV/EBITDA et EV/Sales)."""
    values = {}
    for _, row in comps_df.iterrows():
        if row["Méthode"] == "P/E":
            values[row["Méthode"]] = row["Valeur implicite"] / shares_outstanding
        else:
            values[row["Méthode"]] = (row["Valeur implicite"] - net_debt) / shares_outstanding
    return values


def compute_football_field_data(fifty_two_week_low: float | None, fifty_two_week_high: float | None,
                                 scenario_values: list[float], comps_values: dict[str, float]) -> dict:
    """Construit les fourchettes (low, high) par méthode pour le graphique football field."""
    comps_vals = list(comps_values.values())
    data = {
        "DCF (Bear-Bull)": (min(scenario_values), max(scenario_values)),
        "Comparables (min-max)": (min(comps_vals), max(comps_vals)),
    }
    if fifty_two_week_low and fifty_two_week_high:
        data["52 semaines (marché)"] = (fifty_two_week_low, fifty_two_week_high)
    return data


def compute_peer_implied_ebitda_margin(peers_df: pd.DataFrame) -> float | None:
    """
    Marge d'EBITDA implicite moyenne des pairs, dérivée du rapport entre leurs
    multiples (EV/Sales ÷ EV/EBITDA = EBITDA/Sales) plutôt que d'une donnée de
    marge directement disponible. Sert à documenter si la cible se paie une
    prime ou une décote de multiple cohérente avec son propre profil de marge.
    """
    valid = peers_df.dropna(subset=["ev_to_ebitda", "ev_to_revenue"])
    valid = valid[valid["ev_to_ebitda"] != 0]
    if valid.empty:
        return None
    implied_margins = valid["ev_to_revenue"] / valid["ev_to_ebitda"]
    return float(implied_margins.mean())


def build_valuation_summary(dcf_scenarios: dict, comps_values: dict, consensus: dict,
                             week52_low: float | None, week52_high: float | None,
                             current_price: float | None) -> dict:
    """
    Construit la fourchette de valorisation (football field) et un prix cible
    en conservant TOUTES les méthodes disponibles — DCF (Bear/Base/Bull),
    comparables (EV/EBITDA, EV/Sales, P/E) et consensus analystes — sans
    filtrage mécanique d'aucune d'entre elles.

    Le prix cible est la moyenne simple du DCF (scénario Base), de la moyenne
    des comparables et du consensus analystes, quand ils sont disponibles.
    Un écart important entre le DCF (valeur intrinsèque par les flux) et les
    comparables (réalité du marché) n'est pas éliminé : c'est un signal à
    interpréter et à expliquer qualitativement dans la note de synthèse
    (prime de leadership, structure de capital, biais de périmètre des
    pairs...), pas une anomalie à corriger en amont du calcul.

    dcf_scenarios : dict {1: bear, 2: base, 3: bull}
    comps_values  : dict {"EV/EBITDA": ..., "EV/Sales": ..., "P/E": ...} (valeur par action)
    """
    dcf_bear, dcf_base, dcf_bull = dcf_scenarios[1], dcf_scenarios[2], dcf_scenarios[3]
    comps_avg = sum(comps_values.values()) / len(comps_values) if comps_values else None

    target_components = {"DCF (Base)": dcf_base}
    if comps_avg is not None:
        target_components["Comparables (moyenne)"] = comps_avg
    if consensus.get("target_mean"):
        target_components["Consensus analystes"] = consensus["target_mean"]
    target_price = sum(target_components.values()) / len(target_components)

    range_points = [dcf_bear, dcf_bull]
    range_points.extend(comps_values.values())
    if consensus.get("target_low") and consensus.get("target_high"):
        range_points.extend([consensus["target_low"], consensus["target_high"]])
    if week52_low and week52_high:
        range_points.extend([week52_low, week52_high])

    upside_vs_price = (target_price - current_price) / current_price if current_price else None
    dcf_vs_comps_gap = ((comps_avg - dcf_base) / dcf_base) if comps_avg else None
    dcf_vs_consensus_gap = (
        (consensus["target_mean"] - dcf_base) / dcf_base if consensus.get("target_mean") else None
    )

    return {
        "target_price": target_price,
        "target_components": target_components,
        "range_low": min(range_points),
        "range_high": max(range_points),
        "dcf_bear": dcf_bear, "dcf_base": dcf_base, "dcf_bull": dcf_bull,
        "comps_avg": comps_avg, "comps_values": comps_values,
        "consensus_mean": consensus.get("target_mean"),
        "upside_vs_price": upside_vs_price,
        "dcf_vs_comps_gap": dcf_vs_comps_gap,
        "dcf_vs_consensus_gap": dcf_vs_consensus_gap,
    }
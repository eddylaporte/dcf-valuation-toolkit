"""
Tests unitaires pour le toolkit de valorisation (voir dcf_comps_toolkit.py
pour la liste des modules).

Ne couvre volontairement que les fonctions pures (calculs), pas les fonctions
d'extraction réseau (extract_financials, extract_peer_multiples, etc.), qui
nécessitent yfinance et une connexion internet — hors du périmètre de tests
unitaires rapides.

Lancer avec : pytest test_dcf_comps_toolkit.py -v
"""

import numpy as np
import pandas as pd
import pytest

from valuation_config import ValuationAssumptions
from data_cleaning import (
    clean_statement, get_row, get_latest_value, compute_historical_ebit_margin, get_ebit_margin_trend,
)
from valuation_engine import (
    compute_wacc,
    project_fcf,
    run_dcf,
    scenario_inputs,
    run_comps,
    comps_per_share,
    compute_football_field_data,
    compute_peer_implied_ebitda_margin,
    build_valuation_summary,
    compute_peer_average_beta,
    resolve_beta,
)


# ----------------------------------------------------------------------
# Nettoyage
# ----------------------------------------------------------------------

def test_clean_statement_sorts_columns_and_drops_empty_rows():
    df = pd.DataFrame({
        "2023": [100, None],
        "2022": [90, None],
    }, index=["Revenue", "EmptyRow"])
    cleaned = clean_statement(df)
    assert list(cleaned.columns) == ["2022", "2023"]
    assert "EmptyRow" not in cleaned.index


def test_clean_statement_handles_empty_input():
    assert clean_statement(pd.DataFrame()).empty
    assert clean_statement(None).empty


def test_get_row_returns_first_matching_candidate():
    df = pd.DataFrame({"2023": [100]}, index=["Total Revenue"])
    row = get_row(df, ["Revenue", "Total Revenue"])
    assert row is not None
    assert row["2023"] == 100


def test_get_row_returns_none_when_no_candidate_matches():
    df = pd.DataFrame({"2023": [100]}, index=["Total Revenue"])
    assert get_row(df, ["Nonexistent Label"]) is None


def test_get_latest_value_returns_last_non_null():
    df = pd.DataFrame({"2022": [100], "2023": [None]}, index=["Revenue"])
    assert get_latest_value(df, ["Revenue"]) == 100


def test_get_latest_value_returns_default_when_missing():
    df = pd.DataFrame({"2023": [100]}, index=["Revenue"])
    assert get_latest_value(df, ["Nonexistent"], default=-1) == -1


# ----------------------------------------------------------------------
# Marge d'EBIT historique
# ----------------------------------------------------------------------

def test_compute_historical_ebit_margin_simple_average_when_unweighted():
    income = pd.DataFrame({
        "2021": [1000, 100],
        "2022": [1100, 121],
        "2023": [1200, 144],
    }, index=["Total Revenue", "Operating Income"])
    margin = compute_historical_ebit_margin(income, lookback_years=3, weighted=False)
    expected = np.mean([100 / 1000, 121 / 1100, 144 / 1200])
    assert margin == pytest.approx(expected)


def test_compute_historical_ebit_margin_respects_lookback_window():
    income = pd.DataFrame({
        "2020": [1000, 50],   # marge 5% — doit être exclue par lookback_years=2
        "2021": [1000, 100],  # marge 10%
        "2022": [1000, 200],  # marge 20%
    }, index=["Total Revenue", "Operating Income"])
    margin = compute_historical_ebit_margin(income, lookback_years=2, weighted=False)
    assert margin == pytest.approx((0.10 + 0.20) / 2)


def test_compute_historical_ebit_margin_weighted_by_default_favors_recent_years():
    # marge en nette hausse : 10%, 20%, 30% -> la moyenne pondérée (poids 1,2,3)
    # doit être strictement supérieure à la moyenne simple (20%)
    income = pd.DataFrame({
        "2021": [1000, 100], "2022": [1000, 200], "2023": [1000, 300],
    }, index=["Total Revenue", "Operating Income"])
    weighted_margin = compute_historical_ebit_margin(income, lookback_years=3)  # weighted=True par défaut
    simple_margin = compute_historical_ebit_margin(income, lookback_years=3, weighted=False)
    assert weighted_margin > simple_margin
    expected_weighted = (0.10 * 1 + 0.20 * 2 + 0.30 * 3) / (1 + 2 + 3)
    assert weighted_margin == pytest.approx(expected_weighted)


def test_compute_historical_ebit_margin_returns_none_when_rows_missing():
    income = pd.DataFrame({"2023": [1000]}, index=["Total Revenue"])
    assert compute_historical_ebit_margin(income) is None


def test_get_ebit_margin_trend_returns_year_labeled_margins():
    income = pd.DataFrame({
        "2021": [1000, 100], "2022": [1100, 121], "2023": [1200, 144],
    }, index=["Total Revenue", "Operating Income"])
    trend = get_ebit_margin_trend(income, lookback_years=3)
    assert [year for year, _ in trend] == ["2021", "2022", "2023"]
    assert trend[-1][1] == pytest.approx(144 / 1200)


def test_get_ebit_margin_trend_empty_when_rows_missing():
    income = pd.DataFrame({"2023": [1000]}, index=["Total Revenue"])
    assert get_ebit_margin_trend(income) == []


# ----------------------------------------------------------------------
# WACC
# ----------------------------------------------------------------------

def test_compute_wacc_all_equity_equals_cost_of_equity():
    assumptions = ValuationAssumptions(risk_free_rate=0.04, market_risk_premium=0.05, tax_rate=0.25)
    wacc, cost_of_equity = compute_wacc(beta=1.0, equity_value=100, debt_value=0, assumptions=assumptions)
    assert cost_of_equity == pytest.approx(0.04 + 1.0 * 0.05)
    assert wacc == pytest.approx(cost_of_equity)


def test_compute_wacc_blends_equity_and_debt():
    assumptions = ValuationAssumptions(
        risk_free_rate=0.04, market_risk_premium=0.05, cost_of_debt=0.03, tax_rate=0.20,
    )
    # 50/50 equity-dette
    wacc, cost_of_equity = compute_wacc(beta=1.2, equity_value=100, debt_value=100, assumptions=assumptions)
    expected_coe = 0.04 + 1.2 * 0.05
    expected_wacc = 0.5 * expected_coe + 0.5 * 0.03 * (1 - 0.20)
    assert cost_of_equity == pytest.approx(expected_coe)
    assert wacc == pytest.approx(expected_wacc)


def test_compute_wacc_handles_zero_total_capital():
    assumptions = ValuationAssumptions(risk_free_rate=0.04, market_risk_premium=0.05)
    wacc, cost_of_equity = compute_wacc(beta=1.0, equity_value=0, debt_value=0, assumptions=assumptions)
    assert wacc == cost_of_equity  # pas de division par zéro


# ----------------------------------------------------------------------
# Résolution du beta (contrôle croisé vs beta moyen des pairs)
# ----------------------------------------------------------------------

def test_compute_peer_average_beta_averages_valid_betas():
    peers_df = pd.DataFrame({"beta": [0.6, 0.8, None]})
    assert compute_peer_average_beta(peers_df) == pytest.approx(0.7)


def test_compute_peer_average_beta_none_when_column_missing():
    peers_df = pd.DataFrame({"trailing_pe": [10.0]})
    assert compute_peer_average_beta(peers_df) is None


def test_resolve_beta_keeps_raw_beta_when_close_to_peer_average():
    assumptions = ValuationAssumptions()
    beta, note = resolve_beta(beta_raw=0.65, peer_avg_beta=0.70, assumptions=assumptions)
    assert beta == pytest.approx(0.65)
    assert "retenu tel quel" in note


def test_resolve_beta_falls_back_to_peer_average_when_diverging_too_much():
    # 0.41 vs pairs à 0.70 -> écart de ~41%, au-dessus du seuil par défaut (35%)
    assumptions = ValuationAssumptions()
    beta, note = resolve_beta(beta_raw=0.41, peer_avg_beta=0.70, assumptions=assumptions)
    assert beta == pytest.approx(0.70)
    assert "beta moyen des pairs retenu" in note


def test_resolve_beta_uses_peer_average_when_out_of_plausible_range():
    assumptions = ValuationAssumptions()
    beta, note = resolve_beta(beta_raw=0.05, peer_avg_beta=0.75, assumptions=assumptions)
    assert beta == pytest.approx(0.75)
    assert "hors plage plausible" in note


def test_resolve_beta_uses_sector_fallback_when_no_peer_average_available():
    assumptions = ValuationAssumptions(beta_fallback=0.95)
    beta, note = resolve_beta(beta_raw=None, peer_avg_beta=None, assumptions=assumptions)
    assert beta == pytest.approx(0.95)
    assert "valeur de repli sectorielle" in note


# ----------------------------------------------------------------------
# Projection des FCF et DCF
# ----------------------------------------------------------------------

def test_project_fcf_compounds_revenue_growth_correctly():
    fcf_df = project_fcf(
        base_revenue=1000, growth_path=[0.10, 0.10], ebit_margin=0.20, tax_rate=0.25,
        capex_pct_revenue=0.05, da_pct_revenue=0.05, nwc_pct_revenue_change=0.0,
    )
    assert fcf_df.loc[0, "Revenus"] == pytest.approx(1100)
    assert fcf_df.loc[1, "Revenus"] == pytest.approx(1210)


def test_project_fcf_fcf_formula_is_consistent():
    fcf_df = project_fcf(
        base_revenue=1000, growth_path=[0.0], ebit_margin=0.20, tax_rate=0.25,
        capex_pct_revenue=0.05, da_pct_revenue=0.05, nwc_pct_revenue_change=0.01,
    )
    row = fcf_df.iloc[0]
    expected_fcf = row["NOPAT"] + row["D&A"] - row["Capex"] - row["Delta_BFR"]
    assert row["FCF"] == pytest.approx(expected_fcf)


def test_run_dcf_matches_manual_calculation_single_period():
    fcf_df = pd.DataFrame({"FCF": [100.0]})
    wacc, terminal_growth = 0.10, 0.02
    result = run_dcf(fcf_df, wacc, terminal_growth, net_debt=50, shares_outstanding=10)

    discount_factor = 1 / (1 + wacc)
    expected_pv_fcf = 100 * discount_factor
    expected_terminal_value = 100 * (1 + terminal_growth) / (wacc - terminal_growth)
    expected_pv_terminal = expected_terminal_value * discount_factor
    expected_ev = expected_pv_fcf + expected_pv_terminal
    expected_equity = expected_ev - 50
    expected_value_per_share = expected_equity / 10

    assert result["enterprise_value"] == pytest.approx(expected_ev)
    assert result["value_per_share"] == pytest.approx(expected_value_per_share)


def test_run_dcf_higher_wacc_reduces_value_per_share():
    fcf_df = pd.DataFrame({"FCF": [100.0, 105.0, 110.0]})
    low_wacc_result = run_dcf(fcf_df, wacc=0.06, terminal_growth=0.02, net_debt=0, shares_outstanding=10)
    high_wacc_result = run_dcf(fcf_df, wacc=0.12, terminal_growth=0.02, net_debt=0, shares_outstanding=10)
    assert high_wacc_result["value_per_share"] < low_wacc_result["value_per_share"]


def test_run_dcf_zero_shares_outstanding_returns_nan():
    fcf_df = pd.DataFrame({"FCF": [100.0]})
    result = run_dcf(fcf_df, wacc=0.10, terminal_growth=0.02, net_debt=0, shares_outstanding=0)
    assert np.isnan(result["value_per_share"])


# ----------------------------------------------------------------------
# Scénarios
# ----------------------------------------------------------------------

def test_scenario_inputs_bull_has_higher_growth_than_bear():
    assumptions = ValuationAssumptions()
    bear_growth, bear_margin, bear_g = scenario_inputs(assumptions, ebit_margin_base=0.15, scenario=1)
    bull_growth, bull_margin, bull_g = scenario_inputs(assumptions, ebit_margin_base=0.15, scenario=3)
    assert sum(bull_growth) > sum(bear_growth)
    assert bull_margin > bear_margin
    assert bull_g > bear_g


def test_scenario_inputs_base_matches_input_margin():
    assumptions = ValuationAssumptions()
    _, base_margin, _ = scenario_inputs(assumptions, ebit_margin_base=0.15, scenario=2)
    assert base_margin == pytest.approx(0.15)


def test_scenario_inputs_rejects_invalid_scenario():
    assumptions = ValuationAssumptions()
    with pytest.raises(ValueError):
        scenario_inputs(assumptions, ebit_margin_base=0.15, scenario=4)


# ----------------------------------------------------------------------
# Comparables
# ----------------------------------------------------------------------

def test_run_comps_uses_median_peer_multiples():
    peers_df = pd.DataFrame({
        "trailing_pe": [10.0, 12.0, 14.0],
        "ev_to_ebitda": [6.0, 7.0, 8.0],
        "ev_to_revenue": [1.0, 1.5, 2.0],
    })
    target_metrics = {"revenue": 1000, "ebitda": 200, "net_income": 100}
    comps_df = run_comps(target_metrics, peers_df)

    ev_ebitda_row = comps_df[comps_df["Méthode"] == "EV/EBITDA"].iloc[0]
    assert ev_ebitda_row["Multiple médian peers"] == pytest.approx(7.0)
    assert ev_ebitda_row["Valeur implicite"] == pytest.approx(7.0 * 200)


def test_comps_per_share_subtracts_net_debt_for_ev_methods_only():
    comps_df = pd.DataFrame({
        "Méthode": ["EV/EBITDA", "P/E"],
        "Valeur implicite": [1000.0, 800.0],
    })
    values = comps_per_share(comps_df, net_debt=100, shares_outstanding=10)
    assert values["EV/EBITDA"] == pytest.approx((1000 - 100) / 10)  # EV -> equity
    assert values["P/E"] == pytest.approx(800 / 10)                  # P/E déjà en equity value


def test_compute_peer_implied_ebitda_margin_derives_from_multiple_ratio():
    peers_df = pd.DataFrame({
        "ev_to_ebitda": [10.0, 8.0],
        "ev_to_revenue": [2.0, 2.0],
    })
    # marge implicite = EV/Sales / EV/EBITDA = 0.20 et 0.25 -> moyenne 0.225
    margin = compute_peer_implied_ebitda_margin(peers_df)
    assert margin == pytest.approx(0.225)


def test_compute_peer_implied_ebitda_margin_none_when_no_valid_rows():
    peers_df = pd.DataFrame({"ev_to_ebitda": [None], "ev_to_revenue": [None]})
    assert compute_peer_implied_ebitda_margin(peers_df) is None


# ----------------------------------------------------------------------
# Fourchette de valorisation (sans exclusion de méthode)
# ----------------------------------------------------------------------

def test_build_valuation_summary_keeps_every_method_no_exclusion():
    dcf_scenarios = {1: 50.0, 2: 80.0, 3: 120.0}
    comps_values = {"EV/EBITDA": 200.0, "EV/Sales": 60.0, "P/E": 70.0}  # écart volontairement large
    consensus = {"target_mean": 90.0, "target_low": 75.0, "target_high": 105.0}
    summary = build_valuation_summary(dcf_scenarios, comps_values, consensus,
                                       week52_low=55.0, week52_high=115.0, current_price=85.0)
    # les 3 composantes du prix cible sont bien présentes, aucune exclue
    assert set(summary["target_components"].keys()) == {"DCF (Base)", "Comparables (moyenne)", "Consensus analystes"}
    comps_avg = sum(comps_values.values()) / 3
    expected_target = (80.0 + comps_avg + 90.0) / 3
    assert summary["target_price"] == pytest.approx(expected_target)
    # la fourchette englobe le point le plus bas (bear=50) et le plus haut (comps EV/EBITDA=200)
    assert summary["range_low"] == pytest.approx(50.0)
    assert summary["range_high"] == pytest.approx(200.0)


def test_build_valuation_summary_computes_dcf_vs_comps_gap():
    dcf_scenarios = {1: 40.0, 2: 100.0, 3: 160.0}
    comps_values = {"EV/EBITDA": 150.0}
    consensus = {}
    summary = build_valuation_summary(dcf_scenarios, comps_values, consensus,
                                       week52_low=None, week52_high=None, current_price=100.0)
    assert summary["dcf_vs_comps_gap"] == pytest.approx((150.0 - 100.0) / 100.0)


# ----------------------------------------------------------------------
# Football field
# ----------------------------------------------------------------------

def test_compute_football_field_data_includes_market_range_when_available():
    data = compute_football_field_data(
        fifty_two_week_low=70, fifty_two_week_high=110,
        scenario_values=[60, 90, 120], comps_values={"EV/EBITDA": 85, "P/E": 95},
    )
    assert data["DCF (Bear-Bull)"] == (60, 120)
    assert data["Comparables (min-max)"] == (85, 95)
    assert data["52 semaines (marché)"] == (70, 110)


def test_compute_football_field_data_omits_market_range_when_missing():
    data = compute_football_field_data(
        fifty_two_week_low=None, fifty_two_week_high=None,
        scenario_values=[60, 90, 120], comps_values={"EV/EBITDA": 85, "P/E": 95},
    )
    assert "52 semaines (marché)" not in data
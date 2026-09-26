"""
DCF & Comparables Valuation Toolkit
====================================
Point d'entrée : orchestre extraction -> nettoyage -> DCF -> comparables ->
export (Excel, Word, JSON) pour une société donnée. Le détail de chaque
étape vit dans son propre module (voir ci-dessous).

Modules :
    valuation_config.py  — hypothèses de valorisation et sources
    data_extraction.py   — extraction des données de marché (yfinance)
    data_cleaning.py      — nettoyage des états financiers
    valuation_engine.py   — WACC, DCF, comparables, recommandation
    excel_report.py       — export du classeur Excel
    word_report.py        — export de la note de synthèse Word
    json_report.py        — export des résultats clés en JSON

Usage :
    python dcf_comps_toolkit.py --ticker TTE --peers XOM CVX SHEL BP E
"""

import argparse
import logging

from valuation_config import ValuationAssumptions
from data_extraction import extract_financials, extract_peer_multiples, extract_analyst_consensus
from data_cleaning import build_clean_dataset, get_latest_value, compute_historical_ebit_margin, get_ebit_margin_trend
from valuation_engine import (
    compute_wacc, project_fcf, run_dcf, run_scenario_dcf, run_comps, comps_per_share,
    compute_football_field_data, compute_peer_implied_ebitda_margin, build_valuation_summary,
    compute_peer_average_beta, resolve_beta,
)
from excel_report import export_to_excel
from word_report import draft_synthesis_note
from json_report import export_results_json
from company_narratives import get_company_narrative

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger(__name__)


def run_valuation(ticker: str, peers: list[str], assumptions: ValuationAssumptions) -> None:
    logger.info("1/6 — Extraction des données pour %s", ticker)
    raw = extract_financials(ticker)
    peers_df = extract_peer_multiples(peers)
    consensus = extract_analyst_consensus(ticker)

    logger.info("2/6 — Nettoyage")
    clean_data = build_clean_dataset(raw)
    income, balance = clean_data["income_stmt"], clean_data["balance_sheet"]

    total_revenue = get_latest_value(income, ["Total Revenue", "TotalRevenue"])
    ebitda = get_latest_value(income, ["EBITDA"])
    net_income = get_latest_value(income, ["Net Income", "NetIncome"])
    total_debt = get_latest_value(balance, ["Total Debt", "TotalDebt"], default=0)
    cash = get_latest_value(balance, ["Cash And Cash Equivalents", "CashAndCashEquivalents"], default=0)
    net_debt = (total_debt or 0) - (cash or 0)

    market_cap = raw.get("market_cap")
    current_price = raw.get("current_price")
    shares_outstanding = raw["shares_outstanding"]
    if market_cap and current_price:
        implied_shares = market_cap / current_price
        if not shares_outstanding or abs(shares_outstanding / implied_shares - 1) > 0.5:
            logger.warning("sharesOutstanding brut incohérent avec market_cap/prix "
                           "(ratio ADR ou donnée mal mappée possible). Utilisation "
                           "de la valeur implicite du marché : %.0f", implied_shares)
        shares_outstanding = implied_shares

    beta_raw = raw.get("beta")
    peer_avg_beta = compute_peer_average_beta(peers_df)
    beta, beta_note = resolve_beta(beta_raw, peer_avg_beta, assumptions)
    logger.info("Beta retenu : %.2f — %s", beta, beta_note)

    narrative = get_company_narrative(ticker)
    if narrative.get("terminal_growth_override") is not None:
        override = narrative["terminal_growth_override"]
        delta_bear = assumptions.terminal_growth_base - assumptions.terminal_growth_bear
        delta_bull = assumptions.terminal_growth_bull - assumptions.terminal_growth_base
        assumptions.terminal_growth_base = override
        assumptions.terminal_growth_bear = override - delta_bear
        assumptions.terminal_growth_bull = override + delta_bull
        logger.info("Croissance terminale différenciée pour %s : Base=%.2f%% (Bear=%.2f%%, Bull=%.2f%%)",
                    ticker, override * 100, assumptions.terminal_growth_bear * 100,
                    assumptions.terminal_growth_bull * 100)
    terminal_growth_rationale = narrative.get("terminal_growth_rationale")

    ebit_margin_note = "Moyenne pondérée des 3 derniers exercices (yfinance), poids croissants vers l'exercice le plus récent"
    ebit_margin_trend = get_ebit_margin_trend(income)
    if ebit_margin_trend:
        logger.info("Marges d'EBIT historiques (%d derniers exercices) : %s",
                    len(ebit_margin_trend), ", ".join(f"{m:.1%}" for _, m in ebit_margin_trend))
    ebit_margin_simple = compute_historical_ebit_margin(income, weighted=False)
    ebit_margin_base = assumptions.ebit_margin_override or compute_historical_ebit_margin(income, weighted=True)
    if ebit_margin_base is None:
        raise ValueError(
            "Impossible de déterminer la marge d'EBIT depuis l'historique : "
            "fournir ValuationAssumptions.ebit_margin_override manuellement."
        )
    if assumptions.ebit_margin_override:
        ebit_margin_note = "Valeur forcée manuellement (ebit_margin_override)"
    logger.info("Marge d'EBIT retenue (Base, pondérée) pour les projections : %.1f%% (moyenne simple : %s)",
                ebit_margin_base * 100, f"{ebit_margin_simple:.1%}" if ebit_margin_simple is not None else "n/a")

    logger.info("3/6 — DCF (scénario Base)")
    wacc, cost_of_equity = compute_wacc(beta, market_cap or 0, total_debt or 0, assumptions)
    logger.info("WACC : %.2f%% | Coût des fonds propres : %.2f%%", wacc * 100, cost_of_equity * 100)

    fcf_df = project_fcf(total_revenue, assumptions.revenue_growth_base, ebit_margin_base,
                          assumptions.tax_rate, assumptions.capex_pct_revenue,
                          assumptions.da_pct_revenue, assumptions.nwc_pct_revenue_change)
    dcf_result = run_dcf(fcf_df, wacc, assumptions.terminal_growth_base, net_debt, shares_outstanding)
    logger.info("Valeur par action (DCF, Base) : %.2f", dcf_result["value_per_share"])

    if current_price:
        ratio = dcf_result["value_per_share"] / current_price
        logger.info("Prix de marché : %.2f | Ratio DCF/marché : %.2fx", current_price, ratio)
        if ratio > 3 or ratio < 0.33:
            logger.warning("Écart > 3x avec le marché : vérifier les données brutes "
                           "avant de considérer le résultat comme fiable.")

    logger.info("4/6 — Scénarios Bear / Base / Bull")
    scenario_results, scenario_values = {}, {}
    for scen_id, scen_name in [(1, "Bear"), (2, "Base"), (3, "Bull")]:
        result = run_scenario_dcf(total_revenue, ebit_margin_base, net_debt, shares_outstanding,
                                   wacc, assumptions, scen_id)
        scenario_results[scen_id] = result
        scenario_values[scen_id] = result["value_per_share"]
        logger.info("  %s : %.2f", scen_name, result["value_per_share"])

    logger.info("5/6 — Comparables")
    comps_df = run_comps({"revenue": total_revenue, "ebitda": ebitda, "net_income": net_income}, peers_df)
    comps_values = comps_per_share(comps_df, net_debt, shares_outstanding)
    comps_avg = sum(comps_values.values()) / len(comps_values) if comps_values else None
    ff_data = compute_football_field_data(
        raw.get("fifty_two_week_low"), raw.get("fifty_two_week_high"),
        list(scenario_values.values()), comps_values,
    )
    peer_implied_ebitda_margin = compute_peer_implied_ebitda_margin(peers_df)
    target_ebitda_margin = (ebitda / total_revenue) if ebitda and total_revenue else None

    valuation_summary = build_valuation_summary(
        scenario_values, comps_values, consensus,
        raw.get("fifty_two_week_low"), raw.get("fifty_two_week_high"), current_price,
    )
    logger.info("Fourchette de valorisation : %.2f — %.2f | Prix cible : %.2f",
                valuation_summary["range_low"], valuation_summary["range_high"],
                valuation_summary["target_price"])

    logger.info("6/6 — Export")
    export_to_excel(
        path=f"{ticker.lower()}_valuation_output.xlsx", ticker=ticker, peers=peers,
        clean_data=clean_data, assumptions=assumptions, beta=beta, market_cap=market_cap,
        total_debt=total_debt, net_debt=net_debt, shares_outstanding=shares_outstanding,
        ebit_margin_base=ebit_margin_base, ebit_margin_note=ebit_margin_note, total_revenue=total_revenue,
        ebitda=ebitda, net_income=net_income, current_price=current_price, peers_df=peers_df,
        wacc_base=wacc, quote_currency=raw["quote_currency"], consensus=consensus,
        scenario_values=scenario_values, ff_data=ff_data,
    )
    draft_synthesis_note(
        ticker, peers, dcf_result, ebit_margin_base, ebit_margin_trend, target_ebitda_margin,
        comps_df, comps_values, peer_implied_ebitda_margin, scenario_values, consensus,
        current_price, valuation_summary, beta_note, terminal_growth_rationale, ebit_margin_simple,
        path=f"{ticker.lower()}_note_synthese.docx",
    )
    export_results_json(
        path=f"{ticker.lower()}_results.json", ticker=ticker, peers=peers,
        current_price=current_price, dcf_result=dcf_result, scenario_values=scenario_values,
        comps_df=comps_df, comps_avg=comps_avg, consensus=consensus, valuation_summary=valuation_summary,
        wacc=wacc, ebit_margin_base=ebit_margin_base, assumptions=assumptions, ff_data=ff_data,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Toolkit de valorisation DCF & Comparables")
    parser.add_argument("--ticker", default="TTE", help="Ticker de la société cible")
    parser.add_argument("--peers", nargs="+", default=["XOM", "CVX", "SHEL", "BP", "E"],
                         help="Liste des tickers utilisés comme comparables")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    assumptions = ValuationAssumptions()
    run_valuation(args.ticker, args.peers, assumptions)


if __name__ == "__main__":
    main()
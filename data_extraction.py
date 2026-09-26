"""
Extraction des données financières et de marché via yfinance : états
financiers, multiples des pairs, consensus analystes, taux de change.
"""

import logging

import pandas as pd
import yfinance as yf

logger = logging.getLogger(__name__)


def get_fx_rate(from_ccy: str, to_ccy: str = "USD") -> float | None:
    """Récupère un taux de change spot via yfinance."""
    if from_ccy == to_ccy:
        return 1.0
    pair = f"{from_ccy}{to_ccy}=X"
    try:
        hist = yf.Ticker(pair).history(period="5d")
        return float(hist["Close"].iloc[-1])
    except Exception as exc:
        logger.warning("Taux de change %s introuvable (%s). À renseigner manuellement.", pair, exc)
        return None


def extract_financials(ticker: str) -> dict:
    """Récupère bilan, compte de résultat, cash-flow et données de marché."""
    tk = yf.Ticker(ticker)
    info = tk.info
    financial_currency = info.get("financialCurrency", "USD")
    quote_currency = info.get("currency", "USD")

    logger.info("Devise des états financiers : %s | Devise de cotation : %s",
                financial_currency, quote_currency)

    data = {
        "income_stmt": tk.financials,
        "balance_sheet": tk.balance_sheet,
        "cashflow": tk.cashflow,
        "shares_outstanding": info.get("sharesOutstanding"),
        "beta": info.get("beta"),
        "market_cap": info.get("marketCap"),
        "current_price": info.get("currentPrice"),
        "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
        "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
        "financial_currency": financial_currency,
        "quote_currency": quote_currency,
    }

    if financial_currency != quote_currency:
        fx = get_fx_rate(financial_currency, quote_currency)
        if fx:
            logger.info("Conversion des états financiers %s -> %s (taux = %.5f)",
                        financial_currency, quote_currency, fx)
            for key in ("income_stmt", "balance_sheet", "cashflow"):
                if data[key] is not None and not data[key].empty:
                    data[key] = data[key] * fx
        else:
            logger.error("Conversion de devise impossible : les montants restent en %s.",
                         financial_currency)

    return data


def extract_peer_multiples(peers: list[str]) -> pd.DataFrame:
    """Récupère les multiples de valorisation (P/E, EV/EBITDA, EV/Sales) et le beta des pairs."""
    rows = []
    for ticker in peers:
        try:
            info = yf.Ticker(ticker).info
            rows.append({
                "ticker": ticker,
                "name": info.get("shortName"),
                "market_cap": info.get("marketCap"),
                "enterprise_value": info.get("enterpriseValue"),
                "trailing_pe": info.get("trailingPE"),
                "forward_pe": info.get("forwardPE"),
                "ev_to_ebitda": info.get("enterpriseToEbitda"),
                "ev_to_revenue": info.get("enterpriseToRevenue"),
                "beta": info.get("beta"),
            })
        except Exception as exc:
            logger.warning("Impossible de récupérer %s: %s", ticker, exc)
    return pd.DataFrame(rows)


def extract_analyst_consensus(ticker: str) -> dict:
    """Récupère le consensus analystes (prix cible, recommandation) via yfinance."""
    info = yf.Ticker(ticker).info
    consensus = {
        "target_mean": info.get("targetMeanPrice"),
        "target_median": info.get("targetMedianPrice"),
        "target_low": info.get("targetLowPrice"),
        "target_high": info.get("targetHighPrice"),
        "num_analysts": info.get("numberOfAnalystOpinions"),
        "recommendation": info.get("recommendationKey"),
    }
    logger.info("Consensus analystes : cible moyenne %.2f (range %.2f-%.2f, %s analystes, reco: %s)",
                consensus["target_mean"] or float("nan"), consensus["target_low"] or float("nan"),
                consensus["target_high"] or float("nan"), consensus["num_analysts"],
                consensus["recommendation"])
    return consensus
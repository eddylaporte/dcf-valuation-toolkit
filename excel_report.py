"""
Export vers un classeur Excel professionnel : styles (thème "banque
d'affaires"), feuille DCF avec sélecteur de scénario en formules live,
sensibilité, comparables, football field, sommaire et états financiers.
"""

import logging
from datetime import datetime

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.chart import BarChart, Reference

from valuation_config import ValuationAssumptions, ASSUMPTION_SOURCES

logger = logging.getLogger(__name__)


NAVY = "1F3864"
INPUT_BLUE = "0000FF"
LINK_GREEN = "008000"
HIGHLIGHT_YELLOW = "FFFF00"

FONT_NAME = "Arial"
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color=NAVY)
SECTION_FONT = Font(name=FONT_NAME, size=11, bold=True)
LABEL_FONT = Font(name=FONT_NAME, size=10)
BOLD_LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True)
INPUT_FONT = Font(name=FONT_NAME, size=10, color=INPUT_BLUE)
FORMULA_FONT = Font(name=FONT_NAME, size=10, color="000000")
LINK_FONT = Font(name=FONT_NAME, size=10, color=LINK_GREEN)
NOTE_FONT = Font(name=FONT_NAME, size=8, italic=True, color="808080")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
RESULT_FILL = PatternFill("solid", fgColor=HIGHLIGHT_YELLOW)
THIN_BOTTOM = Border(bottom=Side(style="thin", color="BFBFBF"))

FMT_CURRENCY_M = '$#,##0;($#,##0);"-"'
FMT_PRICE = '$#,##0.00'
FMT_PERCENT = '0.0%'
FMT_MULTIPLE = '0.0"x"'
FMT_NUMBER = '#,##0.0'


def style_header_row(ws, row: int, col_start: int, col_end: int) -> None:
    for col in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def autofit_columns(ws, min_width: int = 10, max_width: int = 42) -> None:
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None or not hasattr(cell, "column_letter"):
                continue
            length = len(str(cell.value))
            col = cell.column_letter
            widths[col] = max(widths.get(col, 0), length)
    for col, length in widths.items():
        ws.column_dimensions[col].width = min(max(length + 2, min_width), max_width)


def add_note(ws, row: int, col: int, text: str) -> None:
    ws.cell(row=row, column=col, value=text).font = NOTE_FONT


def write_section_banner(ws, row: int, col_start: int, col_end: int, text: str) -> None:
    """Bandeau de section coloré (fusionné) — remplace un simple label en gras pour mieux
    séparer visuellement les blocs d'une feuille."""
    ws.merge_cells(start_row=row, start_column=col_start, end_row=row, end_column=col_end)
    cell = ws.cell(row=row, column=col_start, value=text)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def add_box_border(ws, min_row: int, max_row: int, min_col: int, max_col: int,
                    color: str = "BFBFBF") -> None:
    """Trace un encadré (bordure fine) autour d'un bloc de cellules pour le faire
    ressortir comme une "carte" plutôt qu'une simple suite de lignes."""
    thin = Side(style="thin", color=color)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                top=thin if r == min_row else existing.top,
                bottom=thin if r == max_row else existing.bottom,
                left=thin if c == min_col else existing.left,
                right=thin if c == max_col else existing.right,
            )


# ----------------------------------------------------------------------
# Feuilles — états financiers historiques
# ----------------------------------------------------------------------

def write_statement_sheet(wb: Workbook, name: str, df: pd.DataFrame, currency: str) -> None:
    ws = wb.create_sheet(name)
    ws.cell(row=1, column=1, value=f"{name} — en millions {currency}").font = TITLE_FONT

    header_row = 3
    ws.cell(row=header_row, column=1, value="Poste").font = HEADER_FONT
    years = list(df.columns)
    for j, year in enumerate(years, start=2):
        ws.cell(row=header_row, column=j, value=str(year))
    style_header_row(ws, header_row, 1, len(years) + 1)

    for i, (label, row_data) in enumerate(df.iterrows(), start=header_row + 1):
        ws.cell(row=i, column=1, value=str(label)).font = BOLD_LABEL_FONT
        for j, year in enumerate(years, start=2):
            val = row_data[year]
            cell = ws.cell(row=i, column=j)
            if pd.notna(val):
                cell.value = val / 1e6
            cell.font = FORMULA_FONT
            cell.number_format = FMT_CURRENCY_M
            cell.border = THIN_BOTTOM

    last_row = header_row + len(df)
    add_note(ws, last_row + 2, 1,
             "Source : yfinance — à recouper avec les états financiers officiels (10-K / 20-F / document d'enregistrement universel).")
    ws.freeze_panes = ws.cell(row=header_row + 1, column=2).coordinate
    autofit_columns(ws)


# ----------------------------------------------------------------------
# Feuille — DCF avec sélecteur de scénario (formules live)
# ----------------------------------------------------------------------

def write_dcf_model_sheet(wb: Workbook, ticker: str, assumptions: ValuationAssumptions,
                           beta: float, market_cap: float, total_debt: float, net_debt: float,
                           shares_outstanding: float, ebit_margin_base: float, ebit_margin_note: str,
                           total_revenue: float, current_price: float | None) -> dict:
    """
    Construit un DCF entièrement piloté par formules, avec un sélecteur de
    scénario (cellule B23 : 1=Bear, 2=Base, 3=Bull) qui recalcule toute la
    chaîne via CHOOSE(). Retourne les références de cellules utiles aux
    autres feuilles.
    """
    ws = wb.create_sheet("DCF - Modèle")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value=f"DCF — {ticker}").font = TITLE_FONT
    add_note(ws, 2, 1, "Bleu = hypothèse modifiable  |  Noir = formule  |  Jaune = résultat clé")

    def input_row(row, label, value, fmt, col=2, note=None):
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = INPUT_FONT
        cell.number_format = fmt
        if note:
            add_note(ws, row, col + 1, note)
        return cell

    # --- Hypothèses générales -------------------------------------------------
    write_section_banner(ws, 3, 1, 5, "HYPOTHÈSES")
    input_row(4, "Taux sans risque", assumptions.risk_free_rate, FMT_PERCENT, note=ASSUMPTION_SOURCES["Taux sans risque"])
    input_row(5, "Prime de risque marché", assumptions.market_risk_premium, FMT_PERCENT, note=ASSUMPTION_SOURCES["Prime de risque marché"])
    input_row(6, "Beta", beta, "0.00", note=ASSUMPTION_SOURCES["Beta"])
    input_row(7, "Coût de la dette (avant impôt)", assumptions.cost_of_debt, FMT_PERCENT, note=ASSUMPTION_SOURCES["Coût de la dette"])
    input_row(8, "Taux d'imposition", assumptions.tax_rate, FMT_PERCENT, note=ASSUMPTION_SOURCES["Taux d'imposition"])
    input_row(9, "Market cap (equity, M)", (market_cap or 0) / 1e6, FMT_CURRENCY_M)
    input_row(10, "Dette totale (M)", (total_debt or 0) / 1e6, FMT_CURRENCY_M)
    input_row(11, "Dette nette (M)", net_debt / 1e6, FMT_CURRENCY_M)
    input_row(12, "Actions en circulation (M)", shares_outstanding / 1e6, FMT_NUMBER)
    input_row(13, "Capex (% du CA)", assumptions.capex_pct_revenue, FMT_PERCENT)
    input_row(14, "D&A (% du CA)", assumptions.da_pct_revenue, FMT_PERCENT)
    input_row(15, "Variation du BFR (% du CA)", assumptions.nwc_pct_revenue_change, FMT_PERCENT)
    add_box_border(ws, 3, 15, 1, 2)

    # --- WACC (CAPM) ------------------------------------------------------
    write_section_banner(ws, 17, 1, 2, "WACC (CAPM)")
    ws.cell(row=18, column=1, value="Coût des fonds propres (CAPM)").font = LABEL_FONT
    ws["B18"] = "=B4+B6*B5"
    ws["B18"].font, ws["B18"].number_format = FORMULA_FONT, FMT_PERCENT

    ws.cell(row=19, column=1, value="Poids Equity").font = LABEL_FONT
    ws["B19"] = "=B9/(B9+B10)"
    ws["B19"].font, ws["B19"].number_format = FORMULA_FONT, FMT_PERCENT

    ws.cell(row=20, column=1, value="Poids Dette").font = LABEL_FONT
    ws["B20"] = "=1-B19"
    ws["B20"].font, ws["B20"].number_format = FORMULA_FONT, FMT_PERCENT

    ws.cell(row=21, column=1, value="WACC").font = BOLD_LABEL_FONT
    ws["B21"] = "=B19*B18+B20*B7*(1-B8)"
    ws["B21"].font = Font(name=FONT_NAME, size=11, bold=True)
    ws["B21"].number_format = FMT_PERCENT
    ws["B21"].fill = RESULT_FILL
    add_box_border(ws, 17, 21, 1, 2)

    # --- Sélecteur de scénario (mis en évidence) ---------------------------
    write_section_banner(ws, 23, 1, 5, "SCÉNARIO ACTIF")
    ws.cell(row=24, column=1, value="Sélecteur (1 = Bear, 2 = Base, 3 = Bull)").font = BOLD_LABEL_FONT
    scen_cell = ws.cell(row=24, column=2, value=assumptions.default_scenario)
    scen_cell.font = Font(name=FONT_NAME, size=12, bold=True, color=INPUT_BLUE)
    scen_cell.number_format = "0"
    scen_cell.fill = PatternFill("solid", fgColor="FFC000")
    ws["C24"] = '=CHOOSE(B24,"Bear","Base","Bull")'
    ws["C24"].font = Font(name=FONT_NAME, size=12, bold=True)
    ws["C24"].fill = PatternFill("solid", fgColor="FFC000")
    add_box_border(ws, 23, 24, 1, 3, color="FFC000")

    for j, label in enumerate(["", "Bear", "Base", "Bull", "Retenu (actif)"], start=1):
        cell = ws.cell(row=25, column=j, value=label)
        if label:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

    ws.cell(row=26, column=1, value="Marge d'EBIT").font = LABEL_FONT
    ws["B26"] = ebit_margin_base + assumptions.ebit_margin_delta_bear
    ws["C26"] = ebit_margin_base
    ws["D26"] = ebit_margin_base + assumptions.ebit_margin_delta_bull
    ws["E26"] = "=CHOOSE($B$24,B26,C26,D26)"
    add_note(ws, 26, 6, ebit_margin_note + " (Bear/Bull = delta ± "
             f"{abs(assumptions.ebit_margin_delta_bear):.0%} vs Base)")
    for col in ("B", "C", "D"):
        ws[f"{col}26"].font, ws[f"{col}26"].number_format = INPUT_FONT, FMT_PERCENT
    ws["E26"].font, ws["E26"].number_format = Font(name=FONT_NAME, size=10, bold=True), FMT_PERCENT

    ws.cell(row=27, column=1, value="Croissance perpétuelle (g)").font = LABEL_FONT
    ws["B27"], ws["C27"], ws["D27"] = assumptions.terminal_growth_bear, assumptions.terminal_growth_base, assumptions.terminal_growth_bull
    ws["E27"] = "=CHOOSE($B$24,B27,C27,D27)"
    for col in ("B", "C", "D"):
        ws[f"{col}27"].font, ws[f"{col}27"].number_format = INPUT_FONT, FMT_PERCENT
    ws["E27"].font, ws["E27"].number_format = Font(name=FONT_NAME, size=10, bold=True), FMT_PERCENT
    add_box_border(ws, 25, 27, 1, 5)

    # --- Croissance du CA par scénario, par année --------------------------
    write_section_banner(ws, 29, 1, 6, "CROISSANCE DU CA PAR SCÉNARIO (%)")
    year_cols = ["B", "C", "D", "E", "F"]
    ws.cell(row=30, column=1, value="Année").font = HEADER_FONT
    for i, col in enumerate(year_cols, start=1):
        ws[f"{col}30"] = f"N+{i}"
    style_header_row(ws, 30, 1, 6)

    scenario_rows = {"Bear": (31, assumptions.revenue_growth_bear),
                      "Base": (32, assumptions.revenue_growth_base),
                      "Bull": (33, assumptions.revenue_growth_bull)}
    for label, (row, path) in scenario_rows.items():
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        for col, val in zip(year_cols, path):
            cell = ws[f"{col}{row}"]
            cell.value, cell.font, cell.number_format = val, INPUT_FONT, FMT_PERCENT

    ws.cell(row=34, column=1, value="Retenue (active)").font = BOLD_LABEL_FONT
    for col in year_cols:
        cell = ws[f"{col}34"]
        cell.value = f"=CHOOSE($B$24,{col}31,{col}32,{col}33)"
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.number_format = FMT_PERCENT
    add_box_border(ws, 29, 34, 1, 6)

    # --- Projection des FCF (utilise les lignes "actives") -----------------
    write_section_banner(ws, 36, 1, 6, "PROJECTION DES FLUX DE TRÉSORERIE DISPONIBLES")
    input_row(37, "CA de référence (dernier exercice, M)", total_revenue / 1e6, FMT_CURRENCY_M)

    ws.cell(row=38, column=1, value="Année").font = HEADER_FONT
    for i, col in enumerate(year_cols, start=1):
        ws[f"{col}38"] = f"N+{i}"
    style_header_row(ws, 38, 1, 6)

    labels = {39: "Revenus (M)", 40: "EBIT (M)", 41: "NOPAT (M)", 42: "D&A (M)",
              43: "Capex (M)", 44: "Variation du BFR (M)", 45: "FCF (M)",
              46: "Période (t)", 47: "Facteur d'actualisation", 48: "VA des FCF (M)"}
    band_fill = PatternFill("solid", fgColor="F2F2F2")
    for row, label in labels.items():
        cell = ws.cell(row=row, column=1, value=label)
        cell.font = BOLD_LABEL_FONT if row == 45 else LABEL_FONT
        if row % 2 == 0:
            cell.fill = band_fill

    prev_col = "B37"
    for i, col in enumerate(year_cols, start=1):
        ws[f"{col}39"] = f"={prev_col}*(1+{col}34)"
        ws[f"{col}40"] = f"={col}39*$E$26"
        ws[f"{col}41"] = f"={col}40*(1-$B$8)"
        ws[f"{col}42"] = f"={col}39*$B$14"
        ws[f"{col}43"] = f"={col}39*$B$13"
        ws[f"{col}44"] = f"={col}39*$B$15"
        ws[f"{col}45"] = f"={col}41+{col}42-{col}43-{col}44"
        ws[f"{col}46"] = i
        ws[f"{col}47"] = f"=1/(1+$B$21)^{col}46"
        ws[f"{col}48"] = f"={col}45*{col}47"
        for r in (39, 40, 41, 42, 43, 44, 45, 48):
            c = ws[f"{col}{r}"]
            c.font = BOLD_LABEL_FONT if r == 45 else FORMULA_FONT
            c.number_format = FMT_CURRENCY_M
            if r % 2 == 0:
                c.fill = band_fill
        ws[f"{col}46"].font, ws[f"{col}46"].number_format = FORMULA_FONT, "0"
        ws[f"{col}47"].font, ws[f"{col}47"].number_format = FORMULA_FONT, "0.000"
        if 46 % 2 == 0:
            ws[f"{col}46"].fill = band_fill
        prev_col = f"{col}39"
    add_box_border(ws, 36, 48, 1, 6)

    # --- Valeur terminale ---------------------------------------------------
    write_section_banner(ws, 50, 1, 2, "VALEUR TERMINALE")
    ws.cell(row=51, column=1, value="Valeur terminale (M)").font = LABEL_FONT
    ws["B51"] = "=F45*(1+E27)/(B21-E27)"
    ws["B51"].font, ws["B51"].number_format = FORMULA_FONT, FMT_CURRENCY_M
    ws.cell(row=52, column=1, value="VA de la valeur terminale (M)").font = LABEL_FONT
    ws["B52"] = "=B51*F47"
    ws["B52"].font, ws["B52"].number_format = FORMULA_FONT, FMT_CURRENCY_M
    add_box_border(ws, 50, 52, 1, 2)

    # --- Résultats ------------------------------------------------------
    write_section_banner(ws, 54, 1, 2, "RÉSULTATS")
    results = [
        (55, "Somme des VA des FCF (M)", "=SUM(B48:F48)", FMT_CURRENCY_M, False),
        (56, "VA de la valeur terminale (M)", "=B52", FMT_CURRENCY_M, False),
        (57, "Valeur d'entreprise — VE (M)", "=B55+B56", FMT_CURRENCY_M, True),
        (58, "Dette nette (M)", "=B11", FMT_CURRENCY_M, False),
        (59, "Valeur des capitaux propres (M)", "=B57-B58", FMT_CURRENCY_M, True),
        (60, "Nombre d'actions (M)", "=B12", FMT_NUMBER, False),
    ]
    for row, label, formula, fmt, bold in results:
        ws.cell(row=row, column=1, value=label).font = BOLD_LABEL_FONT if bold else LABEL_FONT
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = Font(name=FONT_NAME, size=10, bold=bold)
        cell.number_format = fmt

    ws.cell(row=61, column=1, value="Valeur par action (DCF)").font = BOLD_LABEL_FONT
    vps_cell = ws.cell(row=61, column=2, value="=B59/B60")
    vps_cell.font = Font(name=FONT_NAME, size=12, bold=True)
    vps_cell.number_format = FMT_PRICE
    vps_cell.fill = RESULT_FILL

    input_row(62, "Prix de marché actuel", current_price or 0, FMT_PRICE)
    ws.cell(row=63, column=1, value="Ratio DCF / marché").font = BOLD_LABEL_FONT
    ratio_cell = ws.cell(row=63, column=2, value="=B61/B62")
    ratio_cell.font = Font(name=FONT_NAME, size=10, bold=True)
    ratio_cell.number_format = FMT_MULTIPLE
    add_box_border(ws, 54, 63, 1, 2)

    ws.column_dimensions["A"].width = 40
    for col in year_cols:
        ws.column_dimensions[col].width = 13

    return {"sheet_name": "DCF - Modèle", "fcf_row": 45, "period_row": 46,
            "net_debt_cell": "$B$11", "shares_cell": "$B$12",
            "value_per_share_cell": "$B$61", "wacc_cell": "$B$21",
            "terminal_growth_active_cell": "$E$27", "ebit_margin_active_cell": "$E$26",
            "current_price_cell": "$B$62", "ratio_cell": "$B$63", "scenario_cell": "$B$24"}


def write_sensitivity_sheet(wb: Workbook, dcf_refs: dict, wacc_base: float,
                             terminal_growth: float) -> None:
    """Sensibilité valeur/action au WACC et à g, calculée par formule (pas de valeurs figées)."""
    ws = wb.create_sheet("DCF - Sensibilité")
    ws.cell(row=1, column=1, value="Sensibilité — Valeur par action (scénario actif)").font = TITLE_FONT
    add_note(ws, 2, 1, "Recalculée par formule à partir de la feuille DCF - Modèle : "
                        "toute modification des FCF ou du scénario actif s'y répercute automatiquement.")

    sheet_ref = f"'{dcf_refs['sheet_name']}'"
    fcf_range = f"{sheet_ref}!$B${dcf_refs['fcf_row']}:$F${dcf_refs['fcf_row']}"
    period_range = f"{sheet_ref}!$B${dcf_refs['period_row']}:$F${dcf_refs['period_row']}"
    last_fcf = f"{sheet_ref}!$F${dcf_refs['fcf_row']}"
    last_period = f"{sheet_ref}!$F${dcf_refs['period_row']}"
    net_debt = f"{sheet_ref}!{dcf_refs['net_debt_cell']}"
    shares = f"{sheet_ref}!{dcf_refs['shares_cell']}"

    wacc_values = np.round(np.arange(max(wacc_base - 0.02, 0.03), wacc_base + 0.025, 0.01), 4)
    growth_values = np.round(np.arange(max(terminal_growth - 0.01, 0.0), terminal_growth + 0.015, 0.005), 4)

    header_row, header_col = 4, 2
    ws.cell(row=header_row - 1, column=1, value="WACC \\ g").font = BOLD_LABEL_FONT
    for j, g in enumerate(growth_values, start=header_col):
        cell = ws.cell(row=header_row - 1, column=j, value=g)
        cell.font = INPUT_FONT
        cell.number_format = FMT_PERCENT
    style_header_row(ws, header_row - 1, 1, header_col + len(growth_values) - 1)

    for i, w in enumerate(wacc_values, start=header_row):
        wcell = ws.cell(row=i, column=1, value=w)
        wcell.font = INPUT_FONT
        wcell.number_format = FMT_PERCENT
        for j, _ in enumerate(growth_values, start=header_col):
            g_col = get_column_letter(j)
            formula = (
                f"=IFERROR((SUMPRODUCT({fcf_range},1/(1+$A{i})^{period_range})"
                f"+({last_fcf}*(1+{g_col}${header_row - 1}))/($A{i}-{g_col}${header_row - 1})"
                f"/(1+$A{i})^{last_period}-{net_debt})/{shares},\"n/a\")"
            )
            cell = ws.cell(row=i, column=j, value=formula)
            cell.font = FORMULA_FONT
            cell.number_format = FMT_PRICE

    last_row = header_row + len(wacc_values) - 1
    last_col = header_col + len(growth_values) - 1
    rule = ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B",
    )
    ws.conditional_formatting.add(
        f"{get_column_letter(header_col)}{header_row}:{get_column_letter(last_col)}{last_row}", rule)

    ws.column_dimensions["A"].width = 12
    for j in range(header_col, last_col + 1):
        ws.column_dimensions[get_column_letter(j)].width = 12


# ----------------------------------------------------------------------
# Feuilles — Comparables (formules live)
# ----------------------------------------------------------------------

def write_peers_sheet(wb: Workbook, peers_df: pd.DataFrame) -> None:
    ws = wb.create_sheet("Comps - Peers")
    ws.cell(row=1, column=1, value="Comparables boursiers").font = TITLE_FONT

    headers = ["Ticker", "Nom", "Market Cap (M)", "EV (M)", "P/E", "Fwd P/E", "EV/EBITDA", "EV/Sales"]
    header_row = 3
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    style_header_row(ws, header_row, 1, len(headers))

    for i, r in enumerate(peers_df.itertuples(index=False), start=header_row + 1):
        ws.cell(row=i, column=1, value=r.ticker).font = BOLD_LABEL_FONT
        ws.cell(row=i, column=2, value=r.name).font = LABEL_FONT
        vals = [
            (r.market_cap / 1e6 if pd.notna(r.market_cap) else None, FMT_CURRENCY_M),
            (r.enterprise_value / 1e6 if pd.notna(r.enterprise_value) else None, FMT_CURRENCY_M),
            (r.trailing_pe, FMT_MULTIPLE),
            (r.forward_pe, FMT_MULTIPLE),
            (r.ev_to_ebitda, FMT_MULTIPLE),
            (r.ev_to_revenue, FMT_MULTIPLE),
        ]
        for j, (val, fmt) in enumerate(vals, start=3):
            cell = ws.cell(row=i, column=j, value=val)
            cell.font = FORMULA_FONT
            cell.number_format = fmt

    add_note(ws, header_row + len(peers_df) + 2, 1, "Source : yfinance, données de marché à la date d'exécution du script.")
    ws.freeze_panes = ws.cell(row=header_row + 1, column=3).coordinate
    autofit_columns(ws)


def write_comps_result_sheet(wb: Workbook, total_revenue: float, ebitda: float, net_income: float,
                              dcf_refs: dict) -> None:
    ws = wb.create_sheet("Comps - Résultat")
    ws.cell(row=1, column=1, value="Valorisation par comparables").font = TITLE_FONT

    ws.cell(row=3, column=1, value="Métriques de la cible (M)").font = SECTION_FONT
    metrics = [("Revenue (M)", total_revenue / 1e6), ("EBITDA (M)", ebitda / 1e6),
               ("Net income (M)", net_income / 1e6)]
    for i, (label, value) in enumerate(metrics, start=4):
        ws.cell(row=i, column=1, value=label).font = LABEL_FONT
        cell = ws.cell(row=i, column=2, value=value)
        cell.font, cell.number_format = INPUT_FONT, FMT_CURRENCY_M

    sheet_ref = f"'{dcf_refs['sheet_name']}'"
    ws.cell(row=8, column=1, value="Dette nette (M)").font = LABEL_FONT
    ws["B8"] = f"={sheet_ref}!{dcf_refs['net_debt_cell']}"
    ws["B8"].font, ws["B8"].number_format = LINK_FONT, FMT_CURRENCY_M
    ws.cell(row=9, column=1, value="Actions en circulation (M)").font = LABEL_FONT
    ws["B9"] = f"={sheet_ref}!{dcf_refs['shares_cell']}"
    ws["B9"].font, ws["B9"].number_format = LINK_FONT, FMT_NUMBER

    header_row = 12
    headers = ["Méthode", "Multiple médian peers", "Valeur implicite (M)", "Valeur par action implicite"]
    for j, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=j, value=h)
    style_header_row(ws, header_row, 1, len(headers))

    rows = [
        ("EV/EBITDA", "=MEDIAN('Comps - Peers'!G:G)", "=B{r}*$B$5", True),
        ("EV/Sales", "=MEDIAN('Comps - Peers'!H:H)", "=B{r}*$B$4", True),
        ("P/E", "=MEDIAN('Comps - Peers'!E:E)", "=B{r}*$B$6", False),
    ]
    for i, (method, mult_formula, val_template, is_ev) in enumerate(rows, start=header_row + 1):
        ws.cell(row=i, column=1, value=method).font = BOLD_LABEL_FONT
        mcell = ws.cell(row=i, column=2, value=mult_formula)
        mcell.font, mcell.number_format = FORMULA_FONT, FMT_MULTIPLE
        vcell = ws.cell(row=i, column=3, value=val_template.format(r=i))
        vcell.font, vcell.number_format = FORMULA_FONT, FMT_CURRENCY_M
        per_share_formula = f"=(C{i}-$B$8)/$B$9" if is_ev else f"=C{i}/$B$9"
        pcell = ws.cell(row=i, column=4, value=per_share_formula)
        pcell.font, pcell.number_format = FORMULA_FONT, FMT_PRICE

    avg_row = header_row + len(rows) + 1
    ws.cell(row=avg_row, column=1, value="Moyenne des méthodes").font = BOLD_LABEL_FONT
    avg_cell = ws.cell(row=avg_row, column=4,
                        value=f"=AVERAGE(D{header_row + 1}:D{header_row + len(rows)})")
    avg_cell.font = Font(name=FONT_NAME, size=10, bold=True)
    avg_cell.number_format = FMT_PRICE
    avg_cell.fill = RESULT_FILL

    autofit_columns(ws)


# ----------------------------------------------------------------------
# Feuille — Football field
# ----------------------------------------------------------------------

def write_football_field_sheet(wb: Workbook, ff_data: dict, current_price: float | None) -> None:
    ws = wb.create_sheet("Football Field")
    ws.cell(row=1, column=1, value="Football Field — Fourchettes de valorisation").font = TITLE_FONT
    add_note(ws, 2, 1, "Snapshot calculé à la génération du script à partir des scénarios "
                        "Bear/Bull du DCF, du min-max des comparables et du range 52 semaines.")

    header_row = 4
    for j, h in enumerate(["Méthode", "Low", "Range", "High"], start=1):
        ws.cell(row=header_row, column=j, value=h)
    style_header_row(ws, header_row, 1, 4)

    methods = list(ff_data.keys())
    row = header_row + 1
    for method in methods:
        low, high = ff_data[method]
        ws.cell(row=row, column=1, value=method).font = BOLD_LABEL_FONT
        ws.cell(row=row, column=2, value=low).number_format = FMT_PRICE
        ws.cell(row=row, column=3, value=f"=D{row}-B{row}").number_format = FMT_PRICE
        ws.cell(row=row, column=4, value=high).number_format = FMT_PRICE
        for c in (2, 3, 4):
            ws.cell(row=row, column=c).font = FORMULA_FONT
        row += 1

    if current_price:
        margin = max(current_price * 0.005, 0.5)
        ws.cell(row=row, column=1, value="Prix de marché actuel").font = BOLD_LABEL_FONT
        ws.cell(row=row, column=2, value=current_price - margin).number_format = FMT_PRICE
        ws.cell(row=row, column=3, value=f"=D{row}-B{row}").number_format = FMT_PRICE
        ws.cell(row=row, column=4, value=current_price + margin).number_format = FMT_PRICE
        for c in (2, 3, 4):
            ws.cell(row=row, column=c).font = FORMULA_FONT
        row += 1

    last_data_row = row - 1

    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Football Field"
    chart.x_axis.title = "Valeur par action"
    chart.legend = None
    chart.height, chart.width = 9, 20

    low_ref = Reference(ws, min_col=2, min_row=header_row, max_row=last_data_row)
    range_ref = Reference(ws, min_col=3, min_row=header_row, max_row=last_data_row)
    cats = Reference(ws, min_col=1, min_row=header_row + 1, max_row=last_data_row)
    chart.add_data(low_ref, titles_from_data=True)
    chart.add_data(range_ref, titles_from_data=True)
    chart.set_categories(cats)

    chart.series[0].graphicalProperties.noFill = True
    chart.series[1].graphicalProperties.solidFill = NAVY

    ws.add_chart(chart, f"F{header_row}")
    autofit_columns(ws)


# ----------------------------------------------------------------------
# Feuille — Sommaire (vue de synthèse)
# ----------------------------------------------------------------------

def write_summary_sheet(wb: Workbook, ticker: str, peers: list[str], current_price: float | None,
                         consensus: dict, scenario_values: dict) -> None:
    ws = wb.create_sheet("Sommaire", 0)
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value=f"Synthèse de valorisation — {ticker}").font = TITLE_FONT
    ws.cell(row=2, column=1, value=f"Généré le {datetime.now().strftime('%d/%m/%Y')}").font = NOTE_FONT

    # --- Bloc Valorisation --------------------------------------------------
    write_section_banner(ws, 4, 1, 2, "VALORISATION")
    rows = [
        ("Prix de marché actuel", "='DCF - Modèle'!B62", FMT_PRICE, False),
        ("Valeur par action — DCF (scénario actif)", "='DCF - Modèle'!B61", FMT_PRICE, True),
        ("Scénario actif", "='DCF - Modèle'!C24", None, False),
        ("Valeur par action — Comparables (moyenne)", "='Comps - Résultat'!D16", FMT_PRICE, True),
        ("Ratio DCF / marché", "='DCF - Modèle'!B63", FMT_MULTIPLE, False),
        ("WACC", "='DCF - Modèle'!B21", FMT_PERCENT, False),
        ("Croissance perpétuelle retenue", "='DCF - Modèle'!E27", FMT_PERCENT, False),
        ("Marge d'EBIT retenue", "='DCF - Modèle'!E26", FMT_PERCENT, False),
    ]
    row = 5
    for label, formula, fmt, highlight in rows:
        ws.cell(row=row, column=1, value=label).font = BOLD_LABEL_FONT if highlight else LABEL_FONT
        cell = ws.cell(row=row, column=2, value=formula)
        cell.font = Font(name=FONT_NAME, size=12 if highlight else 10, bold=highlight, color=LINK_GREEN)
        if fmt:
            cell.number_format = fmt
        if highlight:
            cell.fill = RESULT_FILL
        row += 1
    add_box_border(ws, 4, row - 1, 1, 2)
    row += 1

    # --- Bloc Scénarios -------------------------------------------------
    write_section_banner(ws, row, 1, 2, "SCÉNARIOS (SNAPSHOT PYTHON)")
    row += 1
    for label, val in [("Bear", scenario_values.get(1)), ("Base", scenario_values.get(2)),
                        ("Bull", scenario_values.get(3))]:
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        cell = ws.cell(row=row, column=2, value=val)
        cell.font, cell.number_format = FORMULA_FONT, FMT_PRICE
        row += 1
    add_box_border(ws, row - 4, row - 1, 1, 2)
    row += 1

    # --- Bloc Consensus analystes ----------------------------------------
    write_section_banner(ws, row, 1, 2, "CONSENSUS ANALYSTES")
    row += 1
    consensus_rows = [
        ("Prix cible moyen", consensus.get("target_mean"), FMT_PRICE),
        ("Range analystes (low - high)",
         f"{consensus['target_low']:.2f} - {consensus['target_high']:.2f}"
         if consensus.get("target_low") and consensus.get("target_high") else "n/a", None),
        ("Nombre d'analystes", consensus.get("num_analysts"), "0"),
        ("Recommandation", consensus.get("recommendation") or "n/a", None),
    ]
    for label, val, fmt in consensus_rows:
        ws.cell(row=row, column=1, value=label).font = LABEL_FONT
        cell = ws.cell(row=row, column=2, value=val)
        cell.font = FORMULA_FONT
        if fmt:
            cell.number_format = fmt
        row += 1
    add_box_border(ws, row - 4, row - 1, 1, 2)
    row += 1

    # --- Bloc Comparables retenus ------------------------------------------
    write_section_banner(ws, row, 1, 2, "COMPARABLES RETENUS")
    row += 1
    ws.cell(row=row, column=1, value=", ".join(peers)).font = LABEL_FONT
    add_box_border(ws, row - 1, row, 1, 2)
    row += 2

    # --- Légende ----------------------------------------------------------
    write_section_banner(ws, row, 1, 2, "LÉGENDE")
    row += 1
    legend = [("Bleu", "Hypothèse modifiable (input)", INPUT_FONT),
              ("Noir", "Formule calculée", FORMULA_FONT),
              ("Vert", "Lien vers un autre onglet", LINK_FONT),
              ("Jaune", "Résultat clé (surligné)", LABEL_FONT)]
    legend_start = row
    for color_name, desc, font in legend:
        ws.cell(row=row, column=1, value=color_name).font = font
        ws.cell(row=row, column=2, value=desc).font = LABEL_FONT
        row += 1
    add_box_border(ws, legend_start, row - 1, 1, 2)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 24


# ----------------------------------------------------------------------
# Export
# ----------------------------------------------------------------------

def export_to_excel(path: str, ticker: str, peers: list[str], clean_data: dict,
                     assumptions: ValuationAssumptions, beta: float, market_cap: float,
                     total_debt: float, net_debt: float, shares_outstanding: float,
                     ebit_margin_base: float, ebit_margin_note: str, total_revenue: float,
                     ebitda: float, net_income: float, current_price: float | None,
                     peers_df: pd.DataFrame, wacc_base: float, quote_currency: str,
                     consensus: dict, scenario_values: dict, ff_data: dict) -> None:
    wb = Workbook()
    wb.remove(wb.active)  # feuille par défaut vide

    dcf_refs = write_dcf_model_sheet(
        wb, ticker, assumptions, beta, market_cap, total_debt, net_debt,
        shares_outstanding, ebit_margin_base, ebit_margin_note, total_revenue, current_price,
    )
    write_sensitivity_sheet(wb, dcf_refs, wacc_base, assumptions.terminal_growth_base)
    write_peers_sheet(wb, peers_df)
    write_comps_result_sheet(wb, total_revenue, ebitda, net_income, dcf_refs)
    write_football_field_sheet(wb, ff_data, current_price)
    write_statement_sheet(wb, "Compte de résultat", clean_data["income_stmt"], quote_currency)
    write_statement_sheet(wb, "Bilan", clean_data["balance_sheet"], quote_currency)
    write_statement_sheet(wb, "Cash-flow", clean_data["cashflow"], quote_currency)
    write_summary_sheet(wb, ticker, peers, current_price, consensus, scenario_values)  # insérée en 1ère position

    wb.save(path)
    logger.info("Classeur Excel exporté : %s", path)
    logger.info("Ouvre-le dans Excel/LibreOffice pour que les formules se recalculent "
               "(elles n'ont pas de valeur mise en cache tant que le fichier n'a pas été ouvert).")
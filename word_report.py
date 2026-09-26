"""
Génération de la note de synthèse au format Word (.docx) : résumé exécutif
en fourchette + prix cible (aucune méthode exclue), thèse d'investissement à
trois piliers, passerelle EV -> Equity Value explicite, risques sectoriels
spécifiques à la société, et description du football field.
"""

import logging
from datetime import datetime

import pandas as pd

from valuation_config import ASSUMPTION_SOURCES
from company_narratives import get_company_narrative

logger = logging.getLogger(__name__)


def _position_in_range(current_price: float | None, range_low: float, range_high: float) -> str | None:
    """Décrit qualitativement où se situe le cours actuel dans la fourchette de valorisation."""
    if not current_price or range_high <= range_low:
        return None
    position = (current_price - range_low) / (range_high - range_low)
    if position < 0.33:
        return "dans le tiers inférieur de la fourchette de valorisation"
    if position > 0.66:
        return "dans le tiers supérieur de la fourchette de valorisation"
    return "au centre de la fourchette de valorisation"


def draft_synthesis_note(ticker: str, peers: list[str], dcf_result: dict, ebit_margin_base: float,
                          ebit_margin_trend: list[tuple[str, float]], target_ebitda_margin: float | None,
                          comps_df: pd.DataFrame, comps_values: dict[str, float],
                          peer_implied_ebitda_margin: float | None, scenario_results: dict,
                          consensus: dict, current_price: float | None, valuation_summary: dict,
                          beta_note: str, terminal_growth_rationale: str | None,
                          ebit_margin_simple: float | None,
                          path: str = "note_synthese.docx") -> None:
    """Génère une note de synthèse en Word présentant une fourchette de valorisation
    et un prix cible construits sur l'ensemble des méthodes disponibles — aucune
    n'est exclue ; les écarts entre elles sont expliqués qualitativement."""
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    narrative = get_company_narrative(ticker)

    doc.add_heading(f"Note de synthèse — Valorisation {ticker}", level=0)
    subtitle = doc.add_paragraph(f"Date : {datetime.now().strftime('%d/%m/%Y')} — {narrative['sector_label']}")
    subtitle.runs[0].italic = True
    subtitle.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    dcf_value = dcf_result["value_per_share"]
    range_low, range_high = valuation_summary["range_low"], valuation_summary["range_high"]
    target_price = valuation_summary["target_price"]
    position_desc = _position_in_range(current_price, range_low, range_high)

    # ------------------------------------------------------------------
    # 1. Résumé exécutif
    # ------------------------------------------------------------------
    doc.add_heading("1. Résumé exécutif", level=1)
    p = doc.add_paragraph()
    p.add_run("Fourchette de valorisation : ").bold = True
    p.add_run(f"${range_low:.2f} — ${range_high:.2f}")
    p2 = doc.add_paragraph()
    p2.add_run("Prix cible : ").bold = True
    target_run = p2.add_run(f"${target_price:.2f}")
    target_run.bold = True
    if valuation_summary.get("upside_vs_price") is not None:
        p2.add_run(f"  ({valuation_summary['upside_vs_price']:+.1%} vs cours actuel)")
    if current_price:
        p3 = doc.add_paragraph()
        p3.add_run("Cours actuel : ").bold = True
        p3.add_run(f"${current_price:.2f}")
        if position_desc:
            p3.add_run(f" — {position_desc}")

    p4 = doc.add_paragraph()
    p4.add_run("Composantes du prix cible (moyenne simple, aucune méthode exclue) : ").bold = True
    components_text = ", ".join(f"{name} ${val:.2f}" for name, val in valuation_summary["target_components"].items())
    p4.add_run(components_text)

    doc.add_paragraph(
        "Le prix cible ci-dessus moyenne à parts égales le DCF (scénario Base), la moyenne "
        "des comparables et le consensus analystes lorsqu'ils sont disponibles — aucune "
        "méthode n'est écartée sur la base d'un écart mécanique. La fourchette (Bear/Bull "
        "du DCF, bornes des comparables, range du consensus, range 52 semaines) donne "
        "l'amplitude réaliste du titre ; un écart significatif entre valeur intrinsèque "
        "(DCF) et valeur de marché (comparables) est interprété en section 4 — Thèse "
        "d'investissement — plutôt que masqué."
    ).italic = True

    # ------------------------------------------------------------------
    # 2. Méthodologie DCF
    # ------------------------------------------------------------------
    doc.add_heading("2. Méthodologie DCF", level=1)
    doc.add_paragraph("Hypothèses de croissance et de marge détaillées année par année : voir l'onglet "
                       "\"DCF - Modèle\" du classeur Excel.")

    p_beta = doc.add_paragraph()
    p_beta.add_run("Beta retenu : ").bold = True
    p_beta.add_run(beta_note)

    if ebit_margin_trend:
        trend_text = ", ".join(f"{year} : {margin:.1%}" for year, margin in ebit_margin_trend)
        margins = [m for _, m in ebit_margin_trend]
        if len(margins) >= 2 and margins[-1] > margins[0] * 1.03:
            trend_word = "haussière"
        elif len(margins) >= 2 and margins[-1] < margins[0] * 0.97:
            trend_word = "baissière"
        else:
            trend_word = "stable"
        p_margin = doc.add_paragraph()
        p_margin.add_run("Marge d'EBIT retenue : ").bold = True
        margin_run = p_margin.add_run(f"{ebit_margin_base:.1%}")
        margin_run.bold = True
        p_margin.add_run(f" — détail annuel : {trend_text}.")
        doc.add_paragraph(
            f"La marge suit une tendance {trend_word} sur la période observée. Plutôt que de se contenter "
            "de le signaler, la marge retenue ci-dessus est une "
            + ("moyenne pondérée donnant davantage de poids aux exercices récents" if trend_word != "stable"
               else "moyenne pondérée (poids croissants vers l'exercice le plus récent, homogène avec les "
                    "autres dossiers)")
            + (f", contre {ebit_margin_simple:.1%} en moyenne simple — cette pondération corrige le biais "
               f"qu'une moyenne plate introduirait sur une tendance {trend_word}."
               if ebit_margin_simple is not None and abs(ebit_margin_simple - ebit_margin_base) > 0.001
               else " (identique à la moyenne simple ici, la marge étant stable).")
        )
    else:
        doc.add_paragraph(f"Marge d'EBIT retenue : {ebit_margin_base:.1%} (détail annuel non disponible).")

    if terminal_growth_rationale:
        p_g = doc.add_paragraph()
        p_g.add_run("Croissance terminale : ").bold = True
        p_g.add_run(terminal_growth_rationale)

    # ------------------------------------------------------------------
    # 3. Méthodologie Comparables
    # ------------------------------------------------------------------
    doc.add_heading("3. Méthodologie Comparables", level=1)
    doc.add_paragraph(f"Peers retenus : {', '.join(peers)}")
    table = doc.add_table(rows=1, cols=len(comps_df.columns))
    table.style = "Light Grid Accent 1"
    for j, col in enumerate(comps_df.columns):
        header = "Valeur implicite (M)" if col == "Valeur implicite" else str(col)
        table.rows[0].cells[j].text = header
        table.rows[0].cells[j].paragraphs[0].runs[0].bold = True
    for _, row_data in comps_df.iterrows():
        cells = table.add_row().cells
        for j, (col_name, val) in enumerate(row_data.items()):
            if col_name == "Valeur implicite" and isinstance(val, (int, float)):
                cells[j].text = f"{val / 1e6:,.0f}"
            elif isinstance(val, (int, float)):
                cells[j].text = f"{val:,.2f}"
            else:
                cells[j].text = str(val)

    doc.add_paragraph(
        "Passerelle Enterprise Value \u2192 Equity Value : pour les méthodes EV/EBITDA et "
        "EV/Sales, la valeur implicite ci-dessus est une valeur d'entreprise (EV) — la "
        "dette nette est déduite avant division par le nombre d'actions pour obtenir la "
        "valeur par action (voir tableau ci-dessous). La méthode P/E aboutit directement à "
        "une valeur de capitaux propres et n'a pas besoin de cet ajustement. Par "
        "simplification, aucun intérêt minoritaire ni action préférentielle n'est retraité "
        "ici (donnée non disponible de façon fiable via l'extraction automatisée) — à "
        "vérifier manuellement si la structure de capital de la société en comporte."
    )
    per_share_table = doc.add_table(rows=1, cols=2)
    per_share_table.style = "Light Grid Accent 1"
    per_share_table.rows[0].cells[0].text = "Méthode"
    per_share_table.rows[0].cells[1].text = "Valeur par action implicite"
    for cell in per_share_table.rows[0].cells:
        cell.paragraphs[0].runs[0].bold = True
    for method, val in comps_values.items():
        cells = per_share_table.add_row().cells
        cells[0].text = method
        cells[1].text = f"${val:.2f}"

    if peer_implied_ebitda_margin is not None and target_ebitda_margin is not None:
        margin_gap = target_ebitda_margin - peer_implied_ebitda_margin
        if abs(margin_gap) > 0.03:
            sens = "supérieure" if margin_gap > 0 else "inférieure"
            doc.add_paragraph(
                f"Marge d'EBITDA de la cible ({target_ebitda_margin:.1%}) très {sens} à la marge "
                f"d'EBITDA implicite moyenne des pairs ({peer_implied_ebitda_margin:.1%}, dérivée du "
                "rapport EV/Sales ÷ EV/EBITDA de chaque pair). "
                + (
                    "Ceci justifie que la société se paie une prime de multiple par rapport à la "
                    "médiane brute du panel plutôt qu'une application mécanique du multiple médian."
                    if margin_gap > 0 else
                    "Ceci suggère une décote de qualité par rapport au panel — à confirmer par une "
                    "analyse plus fine des causes (mix d'activités, structure de coûts, cycle)."
                )
            )
        else:
            doc.add_paragraph(
                f"Marge d'EBITDA de la cible ({target_ebitda_margin:.1%}) proche de la marge implicite "
                f"moyenne des pairs ({peer_implied_ebitda_margin:.1%}) — le multiple médian du panel "
                "est donc appliqué sans ajustement de qualité particulier."
            )

    # ------------------------------------------------------------------
    # 4. Thèse d'investissement
    # ------------------------------------------------------------------
    doc.add_heading("4. Thèse d'investissement", level=1)

    doc.add_heading("4.1 Catalyseurs de croissance", level=2)
    for item in narrative["growth_catalysts"]:
        doc.add_paragraph(item, style="List Bullet")

    doc.add_heading("4.2 Solidité du modèle économique", level=2)
    doc.add_paragraph(narrative["business_model_strength"])

    doc.add_heading("4.3 Analyse de l'écart DCF vs Multiples", level=2)
    gap = valuation_summary.get("dcf_vs_comps_gap")
    if gap is not None:
        doc.add_paragraph(
            f"Écart entre la moyenne des comparables et le DCF (scénario Base) : {gap:+.1%}. "
            + narrative["dcf_vs_multiples_note"]
        )
    else:
        doc.add_paragraph(narrative["dcf_vs_multiples_note"])

    # ------------------------------------------------------------------
    # 5. Scénarios
    # ------------------------------------------------------------------
    doc.add_heading("5. Scénarios (Bear / Base / Bull)", level=1)
    doc.add_paragraph("Valeur par action selon le scénario retenu pour la croissance du CA, "
                       "la marge d'EBIT et la croissance terminale :")
    scen_table = doc.add_table(rows=1, cols=2)
    scen_table.style = "Light Grid Accent 1"
    scen_table.rows[0].cells[0].text = "Scénario"
    scen_table.rows[0].cells[1].text = "Valeur par action"
    for cell in scen_table.rows[0].cells:
        cell.paragraphs[0].runs[0].bold = True
    for label, value in [("Bear", scenario_results.get(1)), ("Base", scenario_results.get(2)),
                          ("Bull", scenario_results.get(3))]:
        cells = scen_table.add_row().cells
        cells[0].text = label
        cells[1].text = f"${value:.2f}" if value is not None else "n/a"
    bear_val, bull_val = scenario_results.get(1), scenario_results.get(3)
    if bear_val and bull_val and bear_val > 0:
        spread = (bull_val - bear_val) / bear_val
        doc.add_paragraph(
            f"L'écart entre les scénarios Bear et Bull représente {spread:.0%} de la valeur du "
            "scénario Bear. Les deltas Bear/Bull appliqués par défaut sont génériques : à "
            "recalibrer avec une analyse sectorielle spécifique avant de les présenter comme "
            "définitifs (voir section 6 pour les risques propres à cette recalibration)."
        )

    # ------------------------------------------------------------------
    # 6. Comparaison au consensus analystes
    # ------------------------------------------------------------------
    doc.add_heading("6. Comparaison au consensus analystes", level=1)
    if consensus.get("target_mean"):
        doc.add_paragraph(
            f"Le consensus des {consensus.get('num_analysts', 'N/A')} analystes suivant le titre "
            f"donne un prix cible moyen de ${consensus['target_mean']:.2f} "
            f"(range ${consensus.get('target_low', float('nan')):.2f} - "
            f"${consensus.get('target_high', float('nan')):.2f}), pour une recommandation "
            f"globale \"{consensus.get('recommendation', 'n/a')}\"."
        )
        gap_consensus = valuation_summary.get("dcf_vs_consensus_gap")
        if gap_consensus is not None:
            abs_gap = abs(gap_consensus)
            sens = "plus optimiste" if gap_consensus > 0 else "plus prudent"
            if abs_gap <= 0.10:
                doc.add_paragraph(
                    "Le consensus analystes et le DCF Base convergent raisonnablement, ce qui renforce "
                    "la robustesse du niveau de valorisation obtenu par les deux approches indépendantes."
                )
            elif abs_gap <= 0.50:
                doc.add_paragraph(
                    f"Le consensus est sensiblement {sens} que le DCF Base ({gap_consensus:+.1%} d'écart). "
                    "Position retenue : le DCF sert d'ancre pour le prix cible ; le consensus est traité "
                    f"comme un scénario {'haut' if gap_consensus > 0 else 'bas'} plutôt que comme un "
                    "signal à intégrer à parts égales, sauf confirmation d'un catalyseur spécifique "
                    "(guidance révisée, nouveau relais de croissance) qui justifierait de s'en rapprocher."
                )
            else:
                doc.add_paragraph(
                    f"Le consensus est très sensiblement {sens} que le DCF Base ({gap_consensus:+.1%} "
                    "d'écart) — un écart trop large pour être expliqué par la seule prime de qualité "
                    "évoquée en section 4.3. Position retenue : "
                    + (
                        "corriger le WACC vers un beta plus défendable (voir section 2) réduit encore la "
                        "valeur intrinsèque plutôt que de la rapprocher du consensus — l'écart n'est donc "
                        "pas un artefact de modélisation qui se résorberait avec des hypothèses plus "
                        "fines. Le consensus est traité comme un scénario haut nécessitant un catalyseur "
                        "de re-rating confirmé (accélération de croissance internationale, effet de "
                        "portefeuille M&A) avant d'être rapproché du prix cible ; en son absence, le DCF "
                        "reste l'ancre du prix cible et le consensus n'est utilisé que comme borne "
                        "supérieure de la fourchette."
                        if gap_consensus > 0 else
                        "le DCF reste l'ancre du prix cible ; le consensus, nettement plus bas, est "
                        "conservé comme scénario dégradé de la fourchette plutôt qu'écarté, mais ne "
                        "justifie pas à lui seul de réviser le prix cible sans élément factuel "
                        "supplémentaire (résultats, guidance) le corroborant."
                    )
                )
    else:
        doc.add_paragraph("Consensus analystes non disponible pour ce titre via yfinance — "
                           "à compléter manuellement (Bloomberg, Refinitiv, rapport broker).")

    # ------------------------------------------------------------------
    # 7. Football Field
    # ------------------------------------------------------------------
    doc.add_heading("7. Football Field", level=1)
    doc.add_paragraph(
        "Le graphique \"Football Field\" du classeur Excel synthétise l'ensemble des "
        "méthodes sur une même échelle : range 52 semaines (marché), DCF (Bear/Base/Bull), "
        "comparables (EV/EBITDA, EV/Sales, P/E), et prix cible du consensus analystes. "
        "C'est la vue d'ensemble à présenter en comité — chaque barre représente une "
        "méthode indépendante, sans qu'aucune ne soit retirée du visuel."
    )

    # ------------------------------------------------------------------
    # 8. Sources et hypothèses
    # ------------------------------------------------------------------
    doc.add_heading("8. Sources et hypothèses", level=1)
    for label, source in ASSUMPTION_SOURCES.items():
        bullet = doc.add_paragraph(style="List Bullet")
        bullet.add_run(f"{label} : ").bold = True
        bullet.add_run(source)

    # ------------------------------------------------------------------
    # 9. Risques et limites
    # ------------------------------------------------------------------
    doc.add_heading("9. Risques et limites", level=1)
    doc.add_paragraph("Risques sectoriels spécifiques :")
    for bullet in narrative["risk_bullets"]:
        doc.add_paragraph(bullet, style="List Bullet")

    doc.add_paragraph("Limites méthodologiques :")
    methodological_risks = [
        "Sensibilité du DCF au WACC et à la croissance terminale : un écart de quelques "
        "dixièmes de point sur l'une ou l'autre hypothèse peut faire varier la valeur par "
        "action de plusieurs dizaines de pourcents (voir l'onglet Sensibilité du classeur) "
        "— la fourchette de cette note doit être lue comme telle, pas comme un chiffre unique.",
        "Échantillon de comparables limité : un multiple individuel atypique peut encore "
        "peser significativement sur la médiane malgré l'élargissement du panel.",
        "Le beta utilisé peut être une valeur de repli sectorielle (voir alerte dans les "
        "logs d'exécution si c'est le cas) : à vérifier auprès d'une source indépendante "
        "avant de figer le WACC.",
    ]
    for bullet in methodological_risks:
        doc.add_paragraph(bullet, style="List Bullet")

    # ------------------------------------------------------------------
    # 10. Conclusion
    # ------------------------------------------------------------------
    doc.add_heading("10. Conclusion", level=1)
    concl = (
        f"{ticker} se valorise dans une fourchette de ${range_low:.2f} à ${range_high:.2f} "
        f"selon l'ensemble des méthodes retenues, pour un prix cible de ${target_price:.2f}"
        + (f" ({valuation_summary['upside_vs_price']:+.1%} vs cours actuel)."
           if valuation_summary.get("upside_vs_price") is not None else ".")
    )
    if position_desc:
        concl += f" Le cours actuel se situe {position_desc}."
    concl += (
        " Cette conclusion s'appuie sur la thèse d'investissement détaillée en section 4 "
        "(catalyseurs de croissance, solidité du modèle, lecture de l'écart DCF/multiples) "
        "et doit être confrontée aux risques sectoriels de la section 9 avant toute décision."
    )
    doc.add_paragraph(concl)

    doc.save(path)
    logger.info("Brouillon de note de synthèse exporté : %s", path)
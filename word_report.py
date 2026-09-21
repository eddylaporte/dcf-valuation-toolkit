"""
Génération de la note de synthèse au format Word (.docx) : résumé exécutif
avec recommandation argumentée, méthodologie DCF/comparables, scénarios,
comparaison au consensus, sources et conclusion.
"""

import logging
from datetime import datetime

import pandas as pd

from valuation_config import ASSUMPTION_SOURCES

logger = logging.getLogger(__name__)


def draft_synthesis_note(ticker: str, peers: list[str], dcf_result: dict, ebit_margin_base: float,
                          comps_df: pd.DataFrame, scenario_results: dict, consensus: dict,
                          comps_avg: float, current_price: float | None, recommendation: dict,
                          path: str = "note_synthese.docx") -> None:
    """Génère une note de synthèse en Word avec une recommandation argumentée à partir
    des chiffres calculés (pas de section à compléter manuellement)."""
    from docx import Document
    from docx.shared import Pt, RGBColor

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    doc.add_heading(f"Note de synthèse — Valorisation {ticker}", level=0)
    subtitle = doc.add_paragraph(f"Date : {datetime.now().strftime('%d/%m/%Y')}")
    subtitle.runs[0].italic = True
    subtitle.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    dcf_value = dcf_result["value_per_share"]
    upside_dcf = (dcf_value - current_price) / current_price if current_price else None
    upside_comps = (comps_avg - current_price) / current_price if current_price and comps_avg else None
    upside_consensus = ((consensus["target_mean"] - current_price) / current_price
                         if current_price and consensus.get("target_mean") else None)

    doc.add_heading("1. Résumé exécutif", level=1)
    p = doc.add_paragraph()
    p.add_run("Valeur par action (DCF, scénario Base) : ").bold = True
    run = p.add_run(f"${dcf_value:.2f}")
    if upside_dcf is not None:
        p.add_run(f"  ({upside_dcf:+.1%} vs cours actuel)")
    p2 = doc.add_paragraph()
    p2.add_run("Valeur par action (Comparables, moyenne) : ").bold = True
    if comps_avg:
        p2.add_run(f"${comps_avg:.2f}")
        if upside_comps is not None:
            p2.add_run(f"  ({upside_comps:+.1%} vs cours actuel)")
    else:
        p2.add_run("n/a")
    p2b = doc.add_paragraph()
    p2b.add_run("Marge d'EBIT retenue (moyenne historique 3 ans) : ").bold = True
    p2b.add_run(f"{ebit_margin_base:.1%}")
    p3 = doc.add_paragraph()
    p3.add_run("Recommandation : ").bold = True
    reco_run = p3.add_run(recommendation["label"])
    reco_run.bold = True
    if recommendation.get("avg_upside") is not None:
        p3.add_run(f"  (upside moyen implicite : {recommendation['avg_upside']:+.1%})")

    reco_expl = doc.add_paragraph()
    reco_expl.add_run(
        "Cette recommandation prend le DCF comme ancre (c'est la méthode la plus rigoureuse, "
        "fondée sur les flux de trésorerie réels) et moyenne son upside implicite avec celui des "
        "comparables et du consensus analystes — sauf si l'une de ces deux méthodes s'écarte du "
        "DCF de plus de 30 points de pourcentage, auquel cas elle est exclue de la moyenne plutôt "
        "que d'y être mélangée : une méthode bruitée (un multiple de pair faussé par un résultat "
        "exceptionnel, par exemple) ne doit pas pouvoir, à elle seule, inverser le sens de la "
        "recommandation. Le titre est classé en Achat si l'upside moyen ainsi calculé dépasse "
        "+15 %, en Vente s'il est inférieur à -10 %, et en Conserver entre les deux. "
        "C'est un point de départ mécanique et documenté, pas un jugement définitif : il ne "
        "remplace pas une analyse qualitative des catalyseurs et risques propres à la thèse "
        "d'investissement (voir sections 7 et 8)."
    ).italic = True

    if recommendation.get("excluded"):
        excl_p = doc.add_paragraph()
        excl_items = ", ".join(
            f"{method} ({upside:+.1%}, écart de {abs(upside - recommendation['upsides']['DCF']) * 100:.0f} pts vs DCF)"
            for method, upside in recommendation["excluded"].items()
        )
        excl_p.add_run(f"Méthode(s) exclue(s) de la moyenne : {excl_items}.").italic = True

    doc.add_heading("2. Méthodologie DCF", level=1)
    doc.add_paragraph("WACC et hypothèses de croissance : voir l'onglet \"DCF - Modèle\" du classeur Excel.")
    doc.add_paragraph(
        f"La marge d'EBIT retenue ({ebit_margin_base:.1%}) est la moyenne des 3 derniers exercices "
        "réels, et non une hypothèse arbitraire. Point de vigilance : si cette marge suit une "
        "tendance marquée (en hausse ou en baisse) sur la période observée, la moyenne peut être "
        "trop optimiste ou trop conservatrice selon le sens de la tendance — comparer la marge "
        "retenue à la marge du dernier exercice seul permet de vérifier ce biais avant de "
        "présenter le chiffre."
    )

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
        "Point de vigilance à discuter : des comparables issus d'un même secteur peuvent avoir "
        "des profils très différents (mix d'activités, exposition géographique, structure de "
        "capital) qui justifient une décote ou une prime par rapport à la médiane — à documenter "
        "explicitement plutôt que d'appliquer le multiple médian sans ajustement."
    )

    doc.add_heading("4. Scénarios (Bear / Base / Bull)", level=1)
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
            "scénario Bear — cette amplitude reflète surtout la sensibilité mécanique du DCF au "
            "WACC et à la croissance terminale plus qu'une vraie distribution de probabilité des "
            "issues possibles. Les deltas Bear/Bull appliqués par défaut sont génériques : à "
            "recalibrer avec une analyse sectorielle spécifique (scénarios de prix des matières "
            "premières, régulation, cycle de la demande) avant de les présenter comme définitifs."
        )

    doc.add_heading("5. Comparaison au consensus analystes", level=1)
    if consensus.get("target_mean"):
        doc.add_paragraph(
            f"Le consensus des {consensus.get('num_analysts', 'N/A')} analystes suivant le titre "
            f"donne un prix cible moyen de ${consensus['target_mean']:.2f} "
            f"(range ${consensus.get('target_low', float('nan')):.2f} - "
            f"${consensus.get('target_high', float('nan')):.2f}), pour une recommandation "
            f"globale \"{consensus.get('recommendation', 'n/a')}\"."
        )
        if upside_consensus is not None and upside_dcf is not None:
            écart = upside_consensus - upside_dcf
            if abs(écart) > 0.10:
                sens = "plus optimiste" if écart > 0 else "plus prudent"
                doc.add_paragraph(
                    f"Le consensus est sensiblement {sens} que le DCF Base ({upside_consensus:+.1%} "
                    f"contre {upside_dcf:+.1%} d'upside implicite). Un écart de cet ordre suggère que "
                    "les analystes intègrent des éléments que le DCF ne capture pas explicitement ici "
                    "— catalyseurs de croissance (nouveaux projets, M&A), révisions de guidance "
                    "récentes, ou au contraire des risques réglementaires/sectoriels non reflétés "
                    "dans les hypothèses par défaut. À creuser avant de trancher entre les deux."
                )
            else:
                doc.add_paragraph(
                    "Le consensus analystes et le DCF Base convergent raisonnablement (écart "
                    f"d'upside inférieur à 10 points, {upside_consensus:+.1%} contre {upside_dcf:+.1%}), "
                    "ce qui renforce la robustesse du niveau de valorisation obtenu par les deux "
                    "approches indépendantes."
                )
    else:
        doc.add_paragraph("Consensus analystes non disponible pour ce titre via yfinance — "
                           "à compléter manuellement (Bloomberg, Refinitiv, rapport broker).")

    doc.add_heading("6. Analyse de sensibilité", level=1)
    doc.add_paragraph("Voir l'onglet \"DCF - Sensibilité\" et le graphique \"Football Field\" du classeur Excel.")
    if current_price and dcf_result.get("value_per_share"):
        doc.add_paragraph(
            "Le tableau de sensibilité croise le WACC et la croissance terminale : c'est le point "
            "de départ pour évaluer la robustesse de la thèse d'investissement — si la "
            "recommandation ne change de sens que dans un coin extrême de la grille (WACC très bas "
            "ET croissance très haute, par exemple), la conviction peut être plus forte que si elle "
            "bascule au moindre ajustement raisonnable des hypothèses."
        )

    doc.add_heading("7. Sources et hypothèses", level=1)
    for label, source in ASSUMPTION_SOURCES.items():
        bullet = doc.add_paragraph(style="List Bullet")
        bullet.add_run(f"{label} : ").bold = True
        bullet.add_run(source)

    doc.add_heading("8. Risques et limites", level=1)
    risk_bullets = [
        "Sensibilité du DCF au WACC et à la croissance terminale : un écart de quelques dixièmes "
        "de point sur l'une ou l'autre hypothèse peut faire varier la valeur par action de "
        "plusieurs dizaines de pourcents (voir l'onglet Sensibilité) — la conclusion de cette note "
        "doit donc être lue comme une fourchette raisonnable, pas un chiffre précis.",
        "Coûts des intrants et pression sur les marges : une entreprise de consommation courante "
        "reste exposée aux variations de coûts (matières premières agricoles, emballage, énergie, "
        "main-d'œuvre) qu'elle ne peut pas toujours répercuter intégralement sur ses prix de vente "
        "sans affecter les volumes — un arbitrage que la marge d'EBIT moyenne lissée sur 3 ans ne "
        "capture pas explicitement.",
        "Évolution des habitudes de consommation : un changement durable de préférences (santé, "
        "prix, canaux de distribution) peut peser sur la croissance du chiffre d'affaires plus "
        "vite que les hypothèses par défaut du modèle ne le supposent — un risque à évaluer au cas "
        "par cas plutôt qu'à ignorer parce qu'il est difficile à quantifier.",
        "Concentration de l'échantillon de comparables : la médiane des pairs repose sur un nombre "
        "limité de sociétés — un multiple individuel atypique peut encore peser significativement "
        "sur le résultat malgré l'élargissement de l'échantillon (voir la méthode exclue en section 1 "
        "si applicable).",
        "Les scénarios Bear/Bull appliquent des deltas génériques (± delta fixe sur la marge, "
        "courbes de croissance par défaut) — à recalibrer avec une analyse sectorielle spécifique "
        "plutôt que de garder les valeurs par défaut du script.",
        "Le beta utilisé peut être une valeur de repli (voir alerte dans les logs d'exécution si "
        "c'est le cas) : à vérifier auprès d'une source indépendante avant de figer le WACC.",
    ]
    for bullet in risk_bullets:
        doc.add_paragraph(bullet, style="List Bullet")

    doc.add_heading("9. Conclusion", level=1)
    concl_parts = [
        f"Le DCF (scénario Base) valorise {ticker} à ${dcf_value:.2f} par action"
        + (f", soit {upside_dcf:+.1%} par rapport au cours actuel" if upside_dcf is not None else "")
        + ". "
    ]
    if comps_avg:
        concl_parts.append(
            f"L'approche par comparables aboutit à une moyenne de ${comps_avg:.2f}"
            + (f" ({upside_comps:+.1%})" if upside_comps is not None else "")
            + ", "
        )
    if consensus.get("target_mean"):
        concl_parts.append(
            f"et le consensus des analystes se situe à ${consensus['target_mean']:.2f} "
            f"({upside_consensus:+.1%})" if upside_consensus is not None else ""
        )
    concl_parts.append(
        f". Sur la base de la règle d'upside moyen documentée en section 1, la recommandation "
        f"retenue est {recommendation['label']}. "
    )
    if recommendation["label"].startswith("Conserver"):
        concl_parts.append(
            "Les trois méthodes convergent globalement vers une valorisation proche du cours "
            "actuel : il n'y a pas de signal fort de sous- ou de sur-valorisation à ce stade. "
            "Un passage à l'achat supposerait soit une révision à la hausse des hypothèses de "
            "croissance ou de marge sur la base d'éléments concrets (guidance, carnet de "
            "commandes, cycle de prix favorable), soit une baisse du cours qui recrée une marge "
            "de sécurité suffisante."
        )
    elif recommendation["label"].startswith("Achat"):
        concl_parts.append(
            "L'upside moyen implicite est significatif : la thèse d'achat repose sur la "
            "conviction que les hypothèses retenues (ou le prix de marché actuel) sont "
            "correctes — le risque principal serait que ces hypothèses de croissance ou de "
            "marge soient trop optimistes, d'où l'importance de les recouper avec la guidance "
            "de la direction et le consensus sectoriel avant de conclure."
        )
    else:
        concl_parts.append(
            "L'upside moyen implicite est négatif : au cours actuel, les trois méthodes "
            "suggèrent une valorisation de marché supérieure à ce que justifient les flux de "
            "trésorerie projetés et les multiples de comparables. Ceci doit être confronté à un "
            "narratif de marché éventuellement plus optimiste (catalyseurs non capturés par le "
            "DCF) avant de conclure définitivement à une survalorisation."
        )
    doc.add_paragraph("".join(concl_parts))

    doc.save(path)
    logger.info("Brouillon de note de synthèse exporté : %s", path)
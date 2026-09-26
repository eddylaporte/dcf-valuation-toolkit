"""
Contenus qualitatifs spécifiques par société pour la note de synthèse :
catalyseurs de croissance, force du modèle économique, explication de
l'écart DCF vs multiples, et risques sectoriels réels (pas de texte
générique). À enrichir/vérifier à la date de l'analyse — ce sont des points
de départ défendables, pas des faits figés.
"""

COMPANY_NARRATIVES = {
    "KO": {
        "sector_label": "Boissons — modèle de franchise de concentré",
        "terminal_growth_override": 0.025,
        "terminal_growth_rationale": (
            "Croissance terminale retenue à 2,5% (au-dessus de l'hypothèse générique de "
            "2,0%) : exposition structurelle aux marchés émergents, où la consommation par "
            "habitant reste inférieure aux marchés matures, et pricing power démontré sur le "
            "portefeuille de marques — deux éléments qui justifient une croissance long terme "
            "légèrement supérieure à la seule inflation américaine."
        ),
        "growth_catalysts": [
            "Pricing power et gestion active du mix (revenue growth management) sur un "
            "portefeuille de marques à forte notoriété, avec une capacité démontrée à "
            "répercuter l'inflation des intrants sans érosion durable des volumes.",
            "Diversification au-delà du soda classique (eaux, boissons énergétiques, "
            "cafés/thés prêts à boire) et expansion continue sur les marchés émergents, "
            "où la consommation par habitant reste inférieure aux marchés matures.",
            "Exécution commerciale digitale renforcée avec le réseau d'embouteilleurs "
            "partenaires, améliorant la disponibilité produit et l'efficacité promotionnelle.",
        ],
        "business_model_strength": (
            "Le modèle de franchise de concentré est structurellement asset-light : "
            "The Coca-Cola Company possède la marque et la formule, mais délègue "
            "l'essentiel de l'embouteillage et de la distribution — capitalistiques — à "
            "un réseau de partenaires embouteilleurs (dont Coca-Cola Europacific "
            "Partners et Coca-Cola FEMSA sur une partie des marchés). Cette structure "
            "explique la marge d'EBITDA élevée et la forte conversion en free cash flow "
            "du groupe par rapport aux capitaux employés."
        ),
        "dcf_vs_multiples_note": (
            "Les comparables (multiples de marché) intègrent souvent une prime pour la "
            "puissance de marque et la structure capitalistique légère du modèle — un DCF "
            "aux hypothèses de croissance et de marge conservatrices peut donc sous-évaluer "
            "cette qualité. À l'inverse, si le DCF ressort au-dessus des multiples, cela "
            "peut signaler que le marché price un risque (taux, change) supérieur à celui "
            "implicite dans le WACC retenu."
        ),
        "risk_bullets": [
            "Matières premières : le sucre, l'aluminium (canettes) et le PET représentent "
            "une part significative des coûts variables ; une hausse durable de ces cours "
            "pèse sur la marge si elle ne peut être intégralement répercutée sur le prix.",
            "Risque de change (FX) : une forte exposition internationale expose les "
            "résultats consolidés en USD aux effets de conversion, indépendamment de la "
            "performance opérationnelle locale en devise constante.",
            "Risque réglementaire : taxes sur les boissons sucrées (soda tax) dans un "
            "nombre croissant de juridictions, et durcissement des obligations de "
            "recyclage / responsabilité élargie du producteur sur les emballages plastiques.",
        ],
    },
    "MCD": {
        "sector_label": "Restauration rapide — modèle de franchise immobilière",
        "terminal_growth_override": 0.02,
        "terminal_growth_rationale": (
            "Croissance terminale alignée sur l'hypothèse générique (2,0%) plutôt que "
            "majorée : le marché américain, cœur du réseau, est proche de la saturation en "
            "nombre d'unités — la croissance long terme provient davantage des gains de "
            "productivité (digital, mix, franchise) que d'une expansion volumique, ce qui ne "
            "justifie pas de prime de croissance terminale par rapport à l'inflation long terme."
        ),
        "growth_catalysts": [
            "Stratégie \"Accelerating the Arches\" : digital, livraison et programme de "
            "fidélité (application MCD) comme leviers de fréquentation et de panier moyen, "
            "avec un mix de ventes digital en progression structurelle.",
            "Plateformes de valeur et innovation menu pour défendre le trafic dans un "
            "contexte de pouvoir d'achat contraint, sans dégrader durablement le mix prix.",
            "Croissance nette du nombre de restaurants portée majoritairement par des "
            "franchisés à l'international, limitant l'intensité capitalistique pour le groupe.",
        ],
        "business_model_strength": (
            "Un réseau franchisé à plus de 90% génère pour McDonald's des revenus de "
            "loyers et redevances adossés à un patrimoine immobilier propre — des flux "
            "récurrents, de nature quasi-obligataire, avec une intensité capitalistique "
            "bien plus faible que celle d'un opérateur en propre. Cette structure explique "
            "la marge d'EBIT très supérieure (de l'ordre de 45-46%) à celle de chaînes "
            "majoritairement en gestion directe."
        ),
        "dcf_vs_multiples_note": (
            "La marge d'EBIT de McDonald's se situe nettement au-dessus de la médiane d'un "
            "panel de pairs mêlant restaurants intégrés et franchisés (Wendy's, Domino's, "
            "Chipotle, Darden) — ce qui justifie que le titre se paie une prime de multiple "
            "par rapport à la médiane brute du panel plutôt qu'une simple application du "
            "multiple médian sans ajustement de qualité."
        ),
        "risk_bullets": [
            "Inflation des coûts d'exploitation : salaires du personnel de restauration "
            "(pression réglementaire sur les salaires minimums dans plusieurs juridictions) "
            "et coûts des protéines (bœuf, volaille) et matières agricoles.",
            "Risque sur le modèle de franchise : la santé financière des franchisés "
            "conditionne directement la régularité du recouvrement des loyers et "
            "redevances — un réseau de franchisés fragilisé est un risque pour le groupe "
            "même si celui-ci n'opère pas directement les restaurants concernés.",
            "Risque macroéconomique : arbitrage du pouvoir d'achat des consommateurs sur "
            "la restauration rapide en période de pression inflationniste, avec un risque "
            "de \"trade-down\" vers des offres moins margées ou la cuisine à domicile.",
        ],
    },
}

GENERIC_NARRATIVE = {
    "sector_label": "Généraliste",
    "terminal_growth_override": None,
    "terminal_growth_rationale": (
        "Croissance terminale laissée à l'hypothèse générique par défaut (voir "
        "ValuationAssumptions.terminal_growth_base) — aucune différenciation sectorielle "
        "appliquée pour ce ticker. [À compléter] : justifier ou ajuster selon la dynamique "
        "de long terme propre à cette société (maturité du marché, pricing power, exposition "
        "géographique)."
    ),
    "growth_catalysts": [
        "[À compléter] Catalyseurs de croissance spécifiques à l'activité de la société.",
    ],
    "business_model_strength": (
        "[À compléter] Description de la structure du modèle économique (intensité "
        "capitalistique, récurrence des revenus, position concurrentielle)."
    ),
    "dcf_vs_multiples_note": (
        "L'écart entre le DCF et les comparables, s'il est significatif, doit être "
        "documenté qualitativement (position concurrentielle, structure de capital, "
        "biais de périmètre des pairs) plutôt qu'ignoré ou corrigé mécaniquement."
    ),
    "risk_bullets": [
        "[À compléter] Risques sectoriels spécifiques à l'activité de la société.",
    ],
}


def get_company_narrative(ticker: str) -> dict:
    """Retourne le contenu qualitatif spécifique au ticker, ou un gabarit générique à compléter."""
    return COMPANY_NARRATIVES.get(ticker.upper(), GENERIC_NARRATIVE)
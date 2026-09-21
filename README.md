# DCF & Comparables Valuation Toolkit

Outil de valorisation d'actions par flux de trésorerie actualisés (DCF) à
trois scénarios, comparables boursiers et cross-check au consensus des
analystes — développé comme projet personnel dans le cadre d'une recherche
de stage en M&A / Portfolio Management.

**Démo interactive :** _lien à ajouter une fois publié sur GitHub Pages_

## Ce que fait l'outil

1. **Extraction** des états financiers et données de marché via `yfinance`
2. **Nettoyage** et conversion de devise automatique
3. **DCF à trois scénarios** (Bear / Base / Bull), avec calcul du WACC (CAPM)
   et de la marge d'EBIT à partir de la moyenne historique réelle (pas une
   hypothèse arbitraire)
4. **Valorisation par comparables** (médiane EV/EBITDA, EV/Sales, P/E)
5. **Cross-check au consensus analystes**, avec une règle de recommandation
   qui exclut toute méthode divergeant de plus de 30 points de l'upside
   implicite du DCF (voir `valuation_engine.py`)
6. **Export** vers :
   - un classeur Excel avec formules live (le sélecteur de scénario
     recalcule tout automatiquement), sensibilité WACC/croissance, et un
     graphique football field
   - une note de synthèse Word avec recommandation argumentée
   - un fichier JSON des résultats clés

## Installation

```bash
python -m venv .venv
source .venv/bin/activate          # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

Pour lancer les tests :
```bash
pip install -r requirements-dev.txt
pytest test_dcf_comps_toolkit.py -v
```

## Usage

```bash
python dcf_comps_toolkit.py --ticker MCD --peers YUM QSR SBUX WEN DPZ CMG DRI
```

Génère `mcd_valuation_output.xlsx`, `mcd_note_synthese.docx` et
`mcd_results.json` dans le dossier courant.

## Structure du code

| Fichier | Rôle |
|---|---|
| `dcf_comps_toolkit.py` | Point d'entrée — orchestre le run complet |
| `valuation_config.py` | Hypothèses de valorisation et sources documentées |
| `data_extraction.py` | Extraction des données de marché (yfinance) |
| `data_cleaning.py` | Nettoyage des états financiers |
| `valuation_engine.py` | WACC, DCF, comparables, règle de recommandation |
| `excel_report.py` | Export du classeur Excel (formules live) |
| `word_report.py` | Export de la note de synthèse Word |
| `json_report.py` | Export des résultats clés en JSON |
| `test_dcf_comps_toolkit.py` | Tests unitaires (fonctions de calcul pures) |

## Limites connues

- Le beta de marché retourné par `yfinance` est parfois aberrant (hors de la
  plage 0.3–2.5) ; dans ce cas un beta de repli sectoriel est utilisé et
  signalé — à vérifier auprès d'une source indépendante avant publication.
- La médiane des comparables reste sensible à un échantillon de pairs trop
  restreint (un multiple individuel atypique peut dominer le résultat) ;
  6-8 pairs par société donnent des résultats sensiblement plus robustes
  que 3.
- Les deltas Bear/Bull (croissance, marge) sont génériques par défaut — à
  recalibrer avec une analyse sectorielle spécifique.
- Ce n'est pas une recommandation d'investissement ; c'est un exercice de
  portfolio personnel.

## Licence

Projet personnel, libre de réutilisation à but pédagogique.

# Modèle de croissance d'Echo sous Excel — à lire en premier

Ce dossier contient le modèle de croissance du notebook (section 9) en **deux versions Excel**, au format du simulateur du professeur (« IMT - Growth Model Simulator.xlsx »). Les deux répondent à la question du board : **« Echo doit-il généraliser les notifications push agressives ? »**

| Fichier | Pour quoi faire | Mode d'emploi |
|---|---|---|
| `echo_growth_model_v1_simple.xlsx` | **Version simple** : l'équation du professeur (U × (1 − churn) + nouveaux), calée sur Echo. Facile à expliquer à l'oral, approximative. | `README_v1_simple.md` |
| `echo_growth_model_v2_cohort.xlsx` | **Version cohortes** : le modèle du notebook, avec les mêmes chiffres au centime près. À utiliser pour les chiffres du rapport. | `README_v2_cohort.md` |
| `COMPARISON.md` | Comparaison des deux versions avec le notebook et recommandation. | — |

**Pour commencer :**
1. Ouvrez un des deux fichiers.
2. Allez sur l'onglet **Inputs**.
3. Changez le scénario dans la cellule jaune **C4** (Base, Pessimistic, Optimistic ou Custom).
4. Regardez les résultats. Les cellules **bleues sur fond jaune** sont les seules à modifier.

## ⚠️ À savoir : les chiffres de la consigne ne sont plus ceux du notebook actuel

Les chiffres cités dans la consigne viennent de la **première version du notebook** (commit `b88fe6a`), par exemple 13 781 utilisateurs actifs en juin 2027 en Base. Le notebook actuel du groupe donne des chiffres **environ 0,4 % plus bas** (13 722), à cause de deux changements de nettoyage :

1. **Budget pub de mai 2025 (manquant)**, remplacé par la moyenne de la colonne (1 296 €) au lieu d'une interpolation (1 115 €).
   - Le CAC de mai devient 3,04 € alors qu'il vaut 2,61 € tous les autres mois.
   - C'est la cause principale : −1,3 % de budget pub sur 12 mois, −0,3 % d'utilisateurs actifs.
2. **Les 300 lignes de minutes impossibles sont supprimées au lieu d'être imputées** : −0,1 % d'utilisateurs actifs.

Le résumé (section 11) du notebook décrit encore l'ancienne méthode. Aucune conclusion ne change. Les classeurs suivent le **notebook actuel**. L'onglet Checks montre les deux références. Je n'ai rien modifié dans le notebook : à vous de décider si vous revenez à l'interpolation.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `export_from_notebook.py` | Exécute le notebook (dans un dossier temporaire, sans toucher aux fichiers du projet) et exporte toutes les données nécessaires dans `model_data.json`. |
| `build_workbooks.py` + `build_v2.py` | Écrivent les deux classeurs à partir de `model_data.json`. |
| `model_data.json` | Données exportées du notebook actuel : courbes, existants, tendances, scénarios, backtest. |
| `reference_original_run.json` | Chiffres de la version originale du notebook (ceux de la consigne), pour l'onglet Checks. |

## Reconstruire les classeurs (après une modification du notebook)

Dans un terminal, depuis ce dossier :

```bash
pip install openpyxl                 # une seule fois (pandas, numpy, scipy, matplotlib, jinja2 sont déjà ceux du notebook)
python export_from_notebook.py       # ~20 s : rejoue le notebook et met à jour model_data.json
python build_workbooks.py            # ~2 s : réécrit les deux .xlsx
```

Ouvrez ensuite les fichiers dans Excel : ils se recalculent à l'ouverture. Dans LibreOffice, si des 0 s'affichent, appuyez sur Ctrl+Maj+F9.

Pour régénérer `reference_original_run.json`, la commande est écrite dans le fichier lui-même.

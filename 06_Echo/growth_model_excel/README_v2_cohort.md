# Version 2 « COHORTES » — `echo_growth_model_v2_cohort.xlsx`

## 1. Ce que c'est, en 5 lignes

- Le **modèle de croissance du notebook (section 9) recalculé dans Excel**, avec de vraies formules : il donne **exactement** les mêmes chiffres que le notebook (écart 0,00 %, voir l'onglet Checks).
- Les utilisateurs sont rangés par **semaine d'inscription (cohorte)** et en **6 segments** : 3 canaux d'acquisition × activé ou non. Chaque segment a sa propre courbe de rétention.
- Il calcule les **3 scénarios** (Pessimiste = push agressif pour tous, Base, Optimiste = onboarding guidé), un scénario **Custom** que vous réglez, et les **8 lignes « what-if »** du notebook.
- Question du board : **« Echo doit-il généraliser les notifications push agressives ? »** Le modèle répond non : −23 % d'utilisateurs actifs et −16 % de revenu sur 12 mois par rapport à la Base.
- C'est la version « fidèle » : à utiliser pour les chiffres du rapport (voir COMPARISON.md).

## 2. Comment l'ouvrir et quoi regarder en premier

1. Double-cliquez sur le fichier. Si Excel affiche « Mode protégé », cliquez sur **« Activer la modification »**. Le fichier contient environ 31 000 formules : l'ouverture peut prendre quelques secondes.
2. Onglet **Guide** : résumé d'une page.
3. Onglet **Inputs** : le choix du scénario en **C4**, les paramètres (lignes 7 à 15), et les **résultats du scénario choisi en C27 à C31**.
4. Onglet **Scenarios** : les 4 scénarios côte à côte (lignes 5 à 8) et deux graphiques à droite (colonne M) : utilisateurs actifs et heures d'écoute, historique et prévisions.
5. Onglet **WhatIf** : un levier à la fois.

Les autres onglets montrent *comment* c'est calculé : vous n'avez pas besoin de les ouvrir pour utiliser le modèle.

## 3. Ce que vous pouvez modifier (et seulement ça)

**Code couleur** (identique au simulateur du professeur, plus une couleur pour les données importées) :

| Apparence | Signification |
|---|---|
| Texte **bleu sur fond jaune** | Vous pouvez modifier. |
| Texte bleu sans jaune | Valeur d'un scénario tirée du notebook. |
| Texte noir | Formule. |
| Texte vert | Copie d'un autre onglet. |
| Gris foncé sur fond gris clair | **Donnée importée** depuis Python (valeur fixe). En italique : valeur extrapolée. |

**Cellules modifiables :**

| Onglet | Cellule | Ce qu'elle signifie | Valeur Base |
|---|---|---|---|
| Inputs | **C4** | Scénario choisi : liste Base / Pessimistic / Optimistic / Custom. | Base |
| Inputs | **F7** | Custom : multiplicateur du budget pub (1 = tendance, 0 = plus de pub, 1,5 = +50 %). | 1 |
| Inputs | **F8** | Custom : CAC, le coût d'un utilisateur payant (€). | 2,61 |
| Inputs | **F9** | Custom : hausse d'activation, la part de nouveaux utilisateurs en plus qui suivent au moins 3 podcasts (0,10 = +10 points). | 0 |
| Inputs | **F10** | Custom : push agressif pour tout le monde, 1 = oui, 0 = non (liste déroulante). | 0 |
| Inputs | **F11** | Custom : ratio de rétention agressif ÷ normal après la semaine 15. | 0,725 |
| Inputs | **F12** | Custom : ratio de revenu par semaine active agressif ÷ normal. | 1,060 |
| Inputs | **F13** | Custom : facteur φ appliqué chaque semaine aux utilisateurs existants si push agressif. | 0,978 |
| Inputs | **F14** | Custom : nombre de semaines pendant lesquelles les existants s'érodent. | 26 |
| Inputs | **F15** | Custom : ratio de pubs par heure d'écoute agressif ÷ normal. | 1,070 |
| NewUsers | **F10** | Budget pub de mai 2025, absent des données (voir la section 7). | 1 296 € |

Les lignes 11 à 15 n'ont d'effet que si F10 = 1 (push agressif).

**Changer de scénario :** cliquez sur Inputs!C4, puis sur la flèche, et choisissez. Inputs (C27 à C31) et l'onglet Simulation suivent. Scenarios et WhatIf calculent toujours tous les cas en même temps.

**Lire les résultats :**

| Cellule | Résultat |
|---|---|
| Inputs!**C27** | Utilisateurs actifs de la dernière semaine (21 juin 2027). |
| Inputs!**C28** | Revenu publicitaire sur 12 mois (€). |
| Inputs!**C29** | Heures d'écoute moyennes par semaine (North Star Metric). |
| Inputs!**C30** | Budget pub dépensé sur 12 mois (€). |
| Inputs!**C31** | Revenu − dépense pub (€). |
| Scenarios!**B5:H8** | Les mêmes résultats pour les 4 scénarios. |
| WhatIf!**B16:F23** | Les 8 tests « un levier à la fois ». |
| WhatIf!**B28:E30** | La valeur d'un utilisateur activé. |

## 4. Visite onglet par onglet

### Inputs
- **Lignes 7 à 15 : un paramètre par ligne**, avec 4 colonnes de scénarios. G est le scénario choisi : `=INDEX(C7:F7;MATCH($C$4;$C$6:$F$6;0))`, comme chez le professeur. MATCH trouve la position du nom choisi, INDEX prend la valeur.
- **En vert dans Base/Pessimiste/Optimiste :** le CAC (calculé dans NewUsers) et le ratio après 15 semaines (pris dans Curves).
- **C19, facteur de calibration** = 7 457 actifs réels ÷ 7 708 prévus par les courbes = **0,967**. Explication dans l'onglet Existing.
- **C20 à C22 :** CPM (9 € pour 1 000 pubs, fixe), revenu par pub (0,009 €), pubs par heure d'écoute (8,57). Ils servent à convertir le revenu en heures d'écoute.

### Scenarios
- **Lignes 5 à 8 :** utilisateurs actifs en juin 2027, heures d'écoute moyennes, revenu sur 12 mois, budget pub, revenu − dépense, écart avec la Base. Tout est copié des blocs de l'onglet Engine.
- **Tableau à partir de la ligne 11 :** l'historique réel (juillet 2025 à juin 2026, importé) puis les 52 semaines prévues de chaque scénario. Il alimente les deux graphiques.

### WhatIf
- **Lignes 5 à 12 :** les réglages de chaque ligne. En bleu, le levier changé ; en vert, les valeurs reprises d'Inputs.
  - Les lignes « Aggressive » reprennent les paramètres mesurés de la colonne Pessimistic.
  - Elles changent seulement la durée d'érosion des existants : 0, 26 ou 52 semaines.
- **Lignes 16 à 23 :** les résultats (calculés dans Engine).
- **Lignes 28 à 30 : la valeur de l'activation.** C'est la LTV sur 2 ans = somme de la courbe de revenu v des âges 0 à 103 (onglet Curves). Exemple paid social :
  - 1,59 € si l'utilisateur n'est pas activé, 3,33 € s'il est activé ;
  - activer un utilisateur vaut donc **1,74 €** ;
  - un utilisateur payant activé rapporte **1,28 fois son CAC**, alors qu'un non activé ne le rembourse pas.

### Simulation (le calcul du scénario choisi, en détail)
C'est l'équivalent de l'onglet « Model » du professeur, mais par cohortes.
- **C5 à C16 :** les paramètres du scénario choisi (en vert). **K5 à K9 :** les résultats.
- **Lignes 18-19 :** pour chaque segment, sa part des nouveaux utilisateurs et son multiplicateur.
  - Exemple organic non activé : 1 − 48,1 % = 51,9 % des inscrits organiques.
  - Pour paid social, le multiplicateur vaut multiplicateur de budget ÷ CAC = 1 ÷ 2,605 = 0,384 utilisateur par euro.
- **Lignes 23 à 74 : les 52 semaines (t = 0 à 51).** Pour chaque segment, 5 colonnes (par exemple C à G pour organic non activé) :

| Colonne | Formule en clair | Exemple organic non activé |
|---|---|---|
| **New users** | inscriptions de la semaine (NewUsers) × multiplicateur × part du segment | semaine 0 : 97,9 × 1 × 51,9 % = **50,8** |
| **S reversed** | courbe de rétention écrite **à l'envers** : la ligne t contient S(51 − t). Si push agressif : × ratio R(âge) jusqu'à 15 semaines, puis × 0,725. | ligne t = 0 : S(51) = 0,168 ; dernière ligne : S(0) = 0,995 |
| **v reversed** | même chose avec la courbe de revenu v (× 1,06 si push agressif) | |
| **WAU** | `SUMPRODUCT(New users des semaines 0..t ; S reversed des lignes 52−t..52)` | voir ci-dessous |
| **Revenue** | même chose avec v | |

**Pourquoi « à l'envers » ?** Les utilisateurs actifs de la semaine t sont la somme, sur toutes les cohortes j déjà inscrites, de (taille de la cohorte j) × (part encore active à l'âge t − j). En écrivant la courbe à l'envers, la cohorte j se retrouve automatiquement en face de son âge. Exemple pour la semaine t = 1 :
- cohorte de la semaine 0 (50,8 personnes, âge 1) : 50,8 × S(1) = 50,8 × 0,9995 = 50,8 ;
- cohorte de la semaine 1 (71,1 personnes, âge 0) : 71,1 × S(0) = 71,1 × 0,995 = 70,8 ;
- **total = 121,5 utilisateurs actifs** de ce segment (cellule F24).

**Colonnes de fin :**

| Colonne | Contenu | Exemple Base, t = 0 |
|---|---|---|
| AG, AH | Nouveaux utilisateurs : actifs et revenu, somme des 6 segments | 258 actifs |
| AI | Facteur push sur les existants = φ^MIN(t + 1 ; durée) si push agressif, sinon 1. En Pessimiste : 0,978 à t = 0, puis 0,978^26 = 0,57 à partir de t = 25. | 1 |
| AJ | Existants = Existing!C × facteur de calibration × AI | 7 543 × 0,967 = 7 298 |
| AK | Revenu des existants (× 1,06 si push agressif) | |
| **AL** | **Total des utilisateurs actifs** = AG + AJ | **7 556** |
| AM | Total du revenu | 779 € |
| AN | Heures d'écoute = AM ÷ 0,009 ÷ (8,57 × ratio si agressif) | 10 100 h |
| AO | Budget pub de la semaine | 270 € (semaine 0 = 5 jours) |

- **Graphique sous le tableau (ligne 77) :** total, existants et nouveaux.

### NewUsers (les nouveaux utilisateurs, par formules)
- **A. Historique mensuel (lignes 6 à 23)**, importé : inscriptions par canal et budget pub, de janvier 2025 à juin 2026. Formules : LN (logarithme) et CAC du mois = budget ÷ inscrits payants.
- **B. Tendances (lignes 27 à 29)** : `SLOPE` (pente) et `INTERCEPT` (ordonnée à l'origine) de LN(série) en fonction du numéro du mois. C'est une droite de moindres carrés sur le logarithme, soit une croissance exponentielle, exactement comme `numpy.polyfit` dans le notebook.
  - Croissance par mois : organique +4,0 %, cross-promo +5,0 %, budget pub +2,84 %.
  - **C31 : CAC = médiane des CAC mensuels = 2,605 €.**
- **Lignes 34 à 36 :** taux d'activation par canal (importés) : organic 48 %, paid 25 %, cross-promo 46 %.
- **D. Prévision jour par jour (lignes 96 à 457)** : valeur de la tendance du mois ÷ nombre de jours du mois.
  - Exemple : 1er juillet 2026 = mois n° 18, EXP(5,702 + 0,0392 × 18) ÷ 31 = **19,6 inscriptions organiques par jour**.
- **C. Semaines (lignes 40 à 91)** : `SUMIFS` additionne les jours de chaque semaine.
  - La semaine 0 (29 juin) ne contient que 5 jours, du 1er au 5 juillet : 97,9 inscrits organiques, contre 137,1 la semaine suivante.
  - Les inscrits des 29 et 30 juin sont déjà dans les données (onglet Existing).

### Curves (données importées)
- **Pour chaque segment et chaque âge de 0 à 169 semaines :**
  - **S** = part des utilisateurs encore actifs à cet âge ;
  - **v** = revenu publicitaire par utilisateur inscrit à cet âge (€).
  - En italique : valeurs extrapolées au-delà de l'âge observé (environ 75 semaines).
- **Colonne O : le ratio R(âge)** agressif ÷ normal mesuré dans le test A/B, de 0 à 15 semaines : 1,00 au départ, puis 0,725 à 15 semaines.
- **Origine :** fonction `build_curves` du notebook, exportée par `export_from_notebook.py`.

### Existing (données importées + calibration)
- **Colonnes C et D :** contribution des ~20 000 utilisateurs inscrits avant le 1er juillet 2026, semaine par semaine. Elle est calculée en Python avec la fonction `simulate` du notebook, **sans** calibration ni effet push : ces deux corrections sont appliquées dans Simulation.
- **Colonnes E à P :** le détail par segment. La colonne Q vérifie que les segments s'additionnent bien au total.
- **C4 à C6 : le facteur de calibration.**
  - Les courbes prévoient 7 708 actifs pour la semaine du 22 juin 2026, alors que la réalité est de 7 457 : facteur = **0,967**.
  - Raison principale : les utilisateurs du bras agressif du test sont partis plus vite que ne le disent les courbes.
- **Pourquoi importé :** recalculer ~470 cohortes × 6 segments en formules rendrait le fichier énorme, pour exactement le même calcul que pour les nouveaux utilisateurs.

### Engine (technique, ne pas modifier)
- 12 copies du bloc de Simulation : Base (ligne 4), Pessimistic (80), Optimistic (156), Custom (232), puis les 8 what-if (une tous les 76 lignes).
- C'est ce qui permet d'avoir tous les scénarios « vivants » en même temps.

### Backtest (valeurs fixes)
- Résultats du backtest du notebook (9.2) : modèle calé sur les données jusqu'au 28 décembre 2025, prévision de janvier à juin 2026.
  - A : inscriptions prévues ;
  - B : inscriptions réelles.
- **Valeurs fixes**, car les courbes de décembre 2025 ne sont pas dans le classeur. Seules les colonnes d'erreur sont des formules.
- **MAPE :** A = 0,84 % (utilisateurs actifs), 1,29 % (revenu) ; B = 2,06 % et 1,72 %.

### Checks
- Chaque résultat est comparé au notebook actuel (cible : écart d'au plus 0,5 %, statut « OK ») et aux chiffres de votre consigne (colonne G).
- En bas, les contrôles internes : Simulation = bloc Engine du même scénario, segments = total, nouveaux utilisateurs = prévision du notebook.
- La dernière ligne compte les « CHECK » ; elle doit valoir **0**.

## 5. Glossaire

| Terme | Définition simple |
|---|---|
| WAU | Utilisateurs actifs dans la semaine (au moins une ouverture de l'app). |
| Cohorte | Tous les utilisateurs inscrits la même semaine. |
| Segment | Canal d'acquisition (store, publicité paid social, cross-promo podcast) × activé ou non. 6 segments. |
| Activation | Suivre au moins 3 podcasts la première semaine (le « aha moment » du notebook). Les activés restent beaucoup plus. |
| Âge | Nombre de semaines depuis l'inscription (0 = semaine d'inscription). |
| Rétention S(âge) | Part d'une cohorte encore active à cet âge. 1 − rétention ≈ churn cumulé. |
| Churn | Part des actifs perdus chaque semaine. Ici il n'est pas constant : il est déduit des courbes S. |
| Courbe de revenu v(âge) | Revenu pub rapporté par utilisateur inscrit à cet âge. Sa somme sur 2 ans donne la LTV. |
| CAC | Coût d'acquisition = budget pub ÷ nouveaux payants (2,61 €). |
| LTV | Revenu total rapporté par un utilisateur pendant sa vie (ici sur 2 ans, avant coûts). |
| LTV:CAC | Ce que rapporte un utilisateur ÷ ce qu'il coûte. En dessous de 1 : perte. |
| Facteur de calibration | Petit coefficient (0,967) qui recale le point de départ du modèle sur la réalité. |
| φ (phi) | Facteur de perte supplémentaire par semaine pour les utilisateurs existants s'ils reçoivent le push agressif (0,978 = −2,2 % par semaine). |
| Ratio de rétention | Rétention du bras agressif ÷ rétention du bras normal dans le test A/B (0,725 après 15 semaines = −27,5 %). |
| Extrapolation | Prolonger une courbe au-delà de ce qu'on a observé (ici par une décroissance exponentielle). |
| Backtest | Se placer dans le passé, prédire une période connue et comparer. |
| MAPE | Erreur moyenne en %. |
| Biais | Erreur systématique (toujours trop haut ou toujours trop bas). |
| North Star Metric | Heures d'écoute par semaine : la métrique choisie dans le notebook. |
| SUMPRODUCT | Fonction Excel qui multiplie deux listes case par case puis additionne : `SUMPRODUCT({1;2};{3;4}) = 1×3 + 2×4 = 11`. |
| SUMIFS | Somme des cellules qui remplissent une condition (ici : les jours d'une même semaine). |
| INDEX / MATCH | MATCH trouve la position d'une valeur dans une liste, INDEX prend la valeur à une position. |

## 6. Lien avec le notebook (section 9) et avec le simulateur du professeur

| Notebook, section 9 | Dans le classeur |
|---|---|
| `build_curves` (courbes S et v par segment) | Onglet **Curves** (importé). |
| `fit_trends` (tendances, CAC, activation) | **NewUsers** sections A et B (formules SLOPE / INTERCEPT / MEDIAN). |
| `forecast_new` (jours → semaines → segments) | **NewUsers** sections C-D, puis lignes 18-19 de Simulation. |
| `simulate` (somme sur les cohortes) | **Simulation** et **Engine** (SUMPRODUCT). Existants : onglet **Existing** (importé). |
| `CALIB` (calage du départ) | Existing!C6, puis Inputs!C19. |
| `NEW_RATIO`, `REV_RATIO`, `PHI`, `APH_RATIO` | Inputs lignes 11 à 15 et Curves colonne O. |
| `run()` et les 3 scénarios | **Scenarios**. |
| Table what-if 9.4 et `ltv_seg` | **WhatIf**. |
| Backtest 9.2 | **Backtest** (valeurs fixes). |

Ce que la version 2 reprend du simulateur du professeur :
- les onglets Guide / Inputs / Scenarios ;
- les colonnes Base / Pessimistic / Optimistic / Custom + SELECTED avec INDEX/MATCH et une liste déroulante en C4 ;
- le code couleur bleu / noir / vert ;
- le tableau de résultats des scénarios et un graphique.

Ce qui change :
- un pas **hebdomadaire** (52 semaines) au lieu d'un pas mensuel ;
- la **rétention par âge** au lieu d'un churn constant ;
- l'onglet « Model » devient Simulation, par cohortes.

## 7. Hypothèses, limites, et que répondre à « pourquoi ce modèle est-il valide ? »

**Hypothèses (toutes celles du notebook) :**
- Chaque segment garde, dans le futur, la rétention et le revenu observés en 2025-2026. Au-delà d'environ 75 semaines, ils sont extrapolés.
- Les inscriptions et le budget pub suivent leur tendance exponentielle.
- Les canaux gratuits ne coûtent rien.
- Marge de 100 % : pas de données de coûts.
- L'effet du push agressif sur les utilisateurs **existants** n'a **jamais été testé**. On suppose −2,2 % d'activité par semaine pendant 26 semaines. L'onglet WhatIf montre 0 et 52 semaines : la conclusion ne change pas.
- Le facteur de calibration reste constant sur toute l'année.

**Limites :**
- Deux blocs sont importés et non recalculés : les courbes et les existants. On ne peut donc pas « changer la rétention » directement, seulement via les paramètres de scénario.
- Le backtest est figé.
- Le CPM est fixe, car les courbes de revenu ont été calculées à 9 €.
- Le fichier est lourd (31 000 formules).
- Le modèle hérite des choix de nettoyage du notebook, y compris le point ci-dessous.

**⚠️ Écart avec les chiffres de votre consigne (à connaître) :**

Vos chiffres (13 781 utilisateurs actifs en Base, etc.) viennent de la **version originale** du notebook (commit b88fe6a). Le notebook actuel du groupe donne des chiffres environ 0,4 % plus bas (13 722), à cause de deux changements de nettoyage :

1. **Budget pub de mai 2025 (absent des données)**, remplacé par **la moyenne de la colonne (1 296 €)** au lieu d'une interpolation (1 115 €).
   - Le CAC de mai passe à 3,04 € alors qu'il vaut 2,61 € tous les autres mois.
   - La tendance du budget baisse (+2,84 % par mois au lieu de +2,98 %).
   - Effet : −1,3 % de budget pub sur 12 mois et −46 utilisateurs actifs en juin 2027 (Base).
   - C'est la cause principale. Vous pouvez le tester en tapant 1115 dans NewUsers!F10.
   - Dans le notebook, la colonne `spend_eur` reste vide pour mai, et c'est une nouvelle colonne `spend_imputed` qui contient la moyenne.
2. **Les 300 lignes de minutes d'écoute impossibles sont supprimées au lieu d'être imputées.** Ces semaines-là, les utilisateurs paraissent inactifs : la rétention en semaine 1 passe de 100 % à 99,95 %. Effet : −13 utilisateurs actifs (−0,1 %).
3. Le résumé (section 11) du notebook parle encore d'imputation et d'interpolation, alors que le code ne les fait plus.

Ces écarts ne changent aucune conclusion.

**Réponse courte à l'oral :**
« Le modèle est construit sur ce qu'on a mesuré : chaque cohorte suit la rétention réelle de son segment, observée pendant jusqu'à 75 semaines. Nous l'avons validé par un backtest : calé uniquement avec les données de 2025, il prédit janvier-juin 2026 avec 1 à 2 % d'erreur. L'erreur restante est justement la trace du test A/B : les utilisateurs du push agressif sont partis plus vite. Les effets du push agressif viennent du test A/B, qui était valide (pas de déséquilibre entre les groupes). La seule hypothèse non testée, l'effet sur les utilisateurs existants, est couverte par l'analyse de sensibilité : 0, 26 ou 52 semaines d'érosion, la conclusion ne change pas. »

## 8. Dépannage

| Problème | Solution |
|---|---|
| `####` dans une cellule | Colonne trop étroite : double-cliquez sur la bordure droite de la lettre de colonne. |
| Rien ne se recalcule | Excel : Formules > Options de calcul > **Automatique** (ou F9). LibreOffice : Données > Calculer > Recalcul automatique. |
| C'est lent | 31 000 formules. Passez en calcul **manuel** pendant vos modifications (Formules > Options de calcul > Manuel), puis F9 pour recalculer. |
| Pas de flèche en Inputs!C4 ou F10 | Cliquez dans la cellule. Si elle a disparu : Données > Validation des données > Liste, avec la source `Base,Pessimistic,Optimistic,Custom` (C4) ou `0,1` (F10). |
| LibreOffice affiche des 0 ou des valeurs anciennes | Outils > Options > LibreOffice Calc > Formule > « Recalcul au chargement » : **Toujours**, ou Ctrl+Maj+F9. |
| `#REF!` | Une ligne ou une colonne a été supprimée : fermez sans enregistrer ou reconstruisez (README principal). |
| `#NOM?` / `#NAME?` | Formule tapée dans la mauvaise langue. Ctrl+Z. |
| Checks affiche « CHECK » | Normal si vous avez modifié une valeur importée ou NewUsers!F10 : les écarts mesurent alors votre changement. Remettez la valeur d'origine pour revenir à 0. |
| Les dates sont des nombres | Format de cellule > Date. |
| Les graphiques sont vides | Le recalcul n'a pas eu lieu (voir plus haut). |
| J'ai modifié le notebook et je veux mettre le classeur à jour | `python export_from_notebook.py` puis `python build_workbooks.py` (README principal). |

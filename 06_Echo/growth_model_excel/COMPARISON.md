# Comparaison : notebook, version 1 « simple » et version 2 « cohortes »

**Référence = notebook actuel** (version du groupe). Les chiffres de la consigne (version originale du notebook) sont environ 0,4 % plus hauts : voir README.md.

Unités : WAU = utilisateurs actifs dans la dernière semaine (21 juin 2027) ; revenu = revenu publicitaire sur 12 mois ; heures = heures d'écoute moyennes par semaine (North Star Metric).

## 1. Les trois scénarios

| Scénario | Indicateur | Notebook | V1 simple | Écart V1 | V2 cohortes | Écart V2 | Consigne (original) |
|---|---|---|---|---|---|---|---|
| Pessimiste | WAU | 10 528 | 9 857 | **−6,4 %** | 10 528 | 0,0 % | 10 587 |
| Pessimiste | Revenu | 46 936 € | 47 513 € | +1,2 % | 46 936 € | 0,0 % | 47 178 € |
| Pessimiste | Heures | 10 938 | 11 073 | +1,2 % | 10 938 | 0,0 % | 10 995 |
| Base | WAU | 13 722 | 13 490 | −1,7 % | 13 722 | 0,0 % | 13 781 |
| Base | Revenu | 56 006 € | 56 570 € | +1,0 % | 56 006 € | 0,0 % | 56 178 € |
| Base | Heures | 13 962 | 14 102 | +1,0 % | 13 962 | 0,0 % | 14 004 |
| Optimiste | WAU | 14 139 | 13 834 | −2,2 % | 14 139 | 0,0 % | 14 200 |
| Optimiste | Revenu | 56 942 € | 57 603 € | +1,2 % | 56 942 € | 0,0 % | 57 118 € |
| Optimiste | Heures | 14 195 | 14 360 | +1,2 % | 14 195 | 0,0 % | 14 239 |
| Tous | Budget pub 12 mois | 23 298 € | 23 781 € | +2,1 % | 23 298 € | 0,0 % | 23 604 € |

**Les écarts entre scénarios**, qui portent la décision :

| Effet | Notebook / V2 | V1 |
|---|---|---|
| Pessimiste vs Base, WAU | −23,3 % | −26,9 % |
| Pessimiste vs Base, revenu | −16,2 % | −16,0 % |
| Optimiste vs Base, WAU | +3,0 % | +2,5 % |
| Optimiste vs Base, revenu | +1,7 % | +1,8 % |

Les deux versions donnent **la même réponse au board** : ne pas généraliser le push agressif (environ −16 % de revenu sur un an). La V1 exagère un peu la chute du nombre d'utilisateurs.

## 2. Les autres tests

| Test | Notebook / V2 | V1 | Lecture |
|---|---|---|---|
| Erreur moyenne de la V1 sur la trajectoire de 52 semaines (MAPE), Base / Pessimiste / Optimiste | — | 1,2 % / **3,9 %** / 1,5 % | La V1 imite bien la Base, mal le Pessimiste. |
| Backtest janvier-juin 2026, MAPE des WAU, A (inscriptions prévues) / B (réelles) | **0,84 % / 2,06 %** | 5,5 % / 4,3 % | Calée seulement sur l'historique, la V1 prévoit beaucoup moins bien. |
| Budget pub × 1,5 : écart de (revenu − dépense) vs Base | −6 760 € | −4 603 € | La V1 **sous-estime de ~1/3** la perte due à la pub payante. |
| Pub payante arrêtée : écart de (revenu − dépense) vs Base | +13 520 € | +9 206 € | Même biais. |
| WAU en juin 2027 si la V1 est calée uniquement sur l'historique | 13 722 | 11 185 (−18 %) | La V1 n'est pas un modèle de prévision autonome. |
| LTV d'un utilisateur paid social | 2,03 € (LTV:CAC 0,78) | 3,07 € (LTV:CAC 1,18) | La LTV simple donne la mauvaise conclusion sur la pub payante. |
| Nombre de formules | — | 6 242 | V2 : 31 359. |

## 3. Ce que chaque version capte ou perd

**Version 1, simple (une équation, churn constant) :**
- **Capte :**
  - le format exact du cours (le professeur et le jury le reconnaissent) ;
  - l'ordre de grandeur des trois scénarios et de leurs écarts de revenu ;
  - les unit economics, en une ligne chacune ;
  - une traduction parlante : **le push agressif équivaut à passer de 3,4 % à 5,1 % de churn par semaine** (de 13,8 % à 20,2 % par mois).
- **Perd :**
  - l'âge des utilisateurs : les nouveaux partent beaucoup plus que les anciens, ce qui explique l'erreur du backtest ;
  - les différences entre canaux, d'où le biais en faveur de la pub payante et la LTV trompeuse ;
  - la forme temporelle des effets : l'érosion du scénario pessimiste dure 26 semaines dans le notebook mais toute l'année en V1 ;
  - la distinction entre nouveaux et anciens utilisateurs.

**Version 2, cohortes :**
- **Capte :** tout le notebook. Mêmes chiffres, écart 0,00 % sur les 3 scénarios, les 32 valeurs what-if, la calibration, la valeur de l'activation et le backtest. Tous les scénarios restent modifiables en direct.
- **Perd ou cache :**
  - deux blocs importés (les courbes de rétention et la contribution des utilisateurs existants) qu'on ne peut pas recalculer dans Excel ;
  - un backtest figé ;
  - un CPM fixe ;
  - une prise en main plus difficile (11 onglets, 31 000 formules, la formule SUMPRODUCT « à l'envers »).

## 4. Recommandation

- **Pour le rapport : la version 2.**
  - Le rapport doit citer **un seul jeu de chiffres**, et ce sont ceux du notebook. La V2 les reproduit à l'identique.
  - La V2 a un backtest à 1-2 % d'erreur, et elle seule traite correctement la pub payante et l'activation.
  - On peut ajouter une phrase tirée de la V1 (« équivalent à +1,7 point de churn par semaine ») pour l'intuition.
- **Pour la soutenance orale : la version 1 pour expliquer, la version 2 pour répondre.**
  - En 8 minutes, l'équation du professeur et l'onglet Model de la V1 s'expliquent en 30 secondes. Basculer Inputs!C4 de Base à Pessimistic, en démonstration, montre l'effet immédiatement.
  - Mais **annoncez les chiffres de la V2** (−23 % d'utilisateurs actifs, −16 % de revenu) pour ne pas présenter deux jeux de chiffres.
  - Gardez la V2 ouverte pour les questions à froid : « et si les anciens utilisateurs s'érodent un an ? » se trouve dans WhatIf ; « le modèle est-il fiable ? » dans Backtest.
- **Si vous ne gardez qu'une version : la V2.** La V1 dépend de la V2 pour son calage ; seule, elle prévoit mal.

## 5. Faiblesses de mes deux versions (soyons critiques)

1. **La V1 est circulaire.** Son churn est calé sur la sortie de la V2, donc elle ne *valide* pas la V2. Elle la *résume*. Ne la présentez pas comme une confirmation indépendante.
2. **La V1 colle mal au scénario pessimiste** (−6,4 % en fin de période). Un churn constant ne sait pas reproduire « 26 semaines d'érosion puis arrêt ». Un churn en deux périodes corrigerait cela, mais s'éloignerait du format du professeur.
3. **La V1 compte 52 semaines pleines**, alors que le notebook commence le 1er juillet : la première semaine n'a que 5 jours. D'où +2,1 % de budget pub.
4. **La V1 utilise l'ARPU mesuré** (0,1036 €), non calé sur le notebook (0,1030 € par utilisateur actif). D'où environ +1 % de revenu dans tous les scénarios. C'est un choix volontaire : les données plutôt qu'un ajustement de plus.
5. **La V2 a deux boîtes noires** : les onglets Curves et Existing sont calculés en Python. Elle est fidèle *au notebook*, pas *aux données brutes*. Si le notebook a une erreur, la V2 l'a aussi : c'est le cas du budget de mai 2025 imputé par la moyenne.
6. **Les deux versions reposent sur une hypothèse jamais testée** : l'effet du push agressif sur les utilisateurs déjà inscrits (−2,2 % par semaine pendant 26 semaines). L'onglet WhatIf de la V2 montre que la conclusion tient de 0 à 52 semaines, mais le chiffre exact du scénario pessimiste en dépend fortement : 11 857 utilisateurs actifs avec 0 semaine d'érosion, 9 775 avec 52.
7. **Les deux versions ignorent les coûts** (marge de 100 %) et supposent que les canaux gratuits ne coûtent rien.
8. **Le facteur de calibration de la V2 (0,967) est appliqué pendant toute l'année.** Il corrige surtout l'effet des utilisateurs du test A/B, qui disparaîtra peu à peu : c'est une légère sous-estimation probable en fin de période.

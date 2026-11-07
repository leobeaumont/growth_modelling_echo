# Version 1 « SIMPLE » — `echo_growth_model_v1_simple.xlsx`

## 1. Ce que c'est, en 5 lignes

- Un tableur qui reprend **exactement la logique du simulateur du professeur** (« IMT - Growth Model Simulator.xlsx ») : une seule équation, appliquée semaine après semaine.
- Il est **calé sur Echo** : point de départ, budget pub, coût d'acquisition (CAC), inscriptions gratuites, revenu par utilisateur viennent du notebook.
- Il rejoue les **3 scénarios du notebook (section 9.3)** : Pessimiste (push agressif pour tout le monde), Base (on ne change rien), Optimiste (onboarding guidé, +10 points d'activation).
- Il répond à la question du board : **« Echo doit-il généraliser les notifications push agressives ? »** Réponse du modèle : non, car le scénario pessimiste perd environ 27 % d'utilisateurs actifs en un an par rapport à la Base.
- C'est la version « pédagogique » : facile à expliquer à l'oral, mais **approximative** (voir COMPARISON.md).

## 2. Comment l'ouvrir et quoi regarder en premier

1. Double-cliquez sur le fichier. Excel ou LibreOffice Calc l'ouvre. Si Excel affiche un bandeau jaune « Mode protégé », cliquez sur **« Activer la modification »**.
2. Les onglets sont **en bas de l'écran**. Commencez par **Guide** : il résume tout en une page.
3. Allez ensuite dans **Inputs**. En haut, la cellule jaune **C4** contient le nom du scénario choisi (« Base » au départ).
4. Allez dans **Model** pour voir le calcul semaine par semaine du scénario choisi, avec un graphique à droite (colonne Q).
5. Enfin, **Scenarios** montre les trois scénarios côte à côte, avec un graphique en bas (ligne 32).

## 3. Ce que vous pouvez modifier (et seulement ça)

**Règle des couleurs** (la même que chez le professeur) :

| Apparence | Signification |
|---|---|
| Texte **bleu sur fond jaune** | Vous pouvez modifier. |
| Texte bleu sans jaune | Valeur d'un scénario, tirée du notebook. Ne la changez que si vous savez pourquoi. |
| Texte noir | Formule. Ne pas toucher. |
| Texte vert | Valeur copiée d'un autre onglet (formule). |
| Gris foncé sur fond gris clair | Donnée importée depuis Python (valeur fixe). |

**Cellules modifiables :**

| Onglet | Cellule | Ce qu'elle signifie |
|---|---|---|
| Inputs | **C4** | Scénario affiché : liste déroulante Base / Pessimistic / Optimistic / Custom. |
| Inputs | **F7** | Custom : utilisateurs actifs au départ (U₀). |
| Inputs | **F8** | Custom : budget pub de la semaine 1 (€). |
| Inputs | **F9** | Custom : croissance du budget pub par semaine (%). |
| Inputs | **F10** | Custom : CAC, le coût d'un utilisateur payant (€). |
| Inputs | **F11** | Custom : nouveaux utilisateurs gratuits en semaine 1 (store + cross-promo podcast). |
| Inputs | **F12** | Custom : croissance des inscriptions gratuites par semaine (%). |
| Inputs | **F13** | Custom : churn hebdomadaire, la part des actifs perdus chaque semaine. |
| Inputs | **F14** | Custom : taux de parrainage par semaine (0 par défaut). |
| Inputs | **F15** | Custom : ARPU, le revenu pub par utilisateur actif et par semaine (€). |
| Inputs | **F16** | Custom : marge brute (100 % par défaut). |
| Inputs | **F17** | Custom : publicités entendues par heure d'écoute. |
| Inputs | **F18** | Custom : revenu par publicité (0,009 € = 9 € pour 1 000 pubs). |
| Scenarios | **C3** | Taille du changement des tests « un levier à la fois » (20 %). |
| Scenarios | **E3** | Multiplicateur de budget pub du test « dépense × 1,5 ». |
| Scenarios | **G3** | Taux de parrainage du test « programme de parrainage ». |
| Calibration | **F3** | Pas du balayage du churn (0,05 point). |

**Changer de scénario :** cliquez sur Inputs!C4, puis sur la petite flèche à droite de la cellule, et choisissez. Les onglets Model et « Unit economics » (Inputs, lignes 23 à 31) se mettent à jour. L'onglet Scenarios montre toujours les quatre scénarios.

**Tester votre propre scénario :** choisissez « Custom » en C4, puis modifiez la colonne F d'Inputs. Exemple : mettez F8 à 0 (plus de pub payante) et regardez Model!I57.

**Lire les résultats :**

| Cellule | Résultat |
|---|---|
| Model!**I57** | Utilisateurs actifs de la dernière semaine (21 juin 2027). |
| Model!**K58** | Revenu publicitaire sur 12 mois (€). |
| Model!**M58** | Budget pub dépensé sur 12 mois (€). |
| Model!**O59** | Heures d'écoute moyennes par semaine (le North Star Metric). |
| Scenarios, colonnes **BO à BU** (lignes 5 à 8) | Les mêmes résultats pour chaque scénario. |

## 4. Visite onglet par onglet

### Guide
Le résumé : équation, code couleur, liste des onglets, 4 exercices à faire.

### Inputs
- **Lignes 7 à 18 : un paramètre par ligne.** Les colonnes C, D, E, F sont les 4 scénarios, G est le scénario choisi, H la source de chaque chiffre. Chaque valeur a aussi un commentaire : passez la souris sur la cellule.
- **Colonne G (SELECTED)** : `=INDEX(C7:F7;MATCH($C$4;$C$6:$F$6;0))`.
  - En clair : MATCH cherche le nom choisi en C4 parmi les titres C6:F6 et renvoie sa position (« Base » donne 1). INDEX prend la valeur à cette position dans C7:F7.
  - Exemple : C4 = « Pessimistic » donne la position 2, donc G13 prend la valeur de D13, soit 5,073 % de churn.
  - (Dans LibreOffice en français, le séparateur des arguments est « ; » au lieu de « , ». C'est la même formule.)
- **Lignes 23 à 31 : unit economics du scénario choisi**, calculées comme chez le professeur. Exemple en Base :
  - Nouveaux payants en semaine 1 = 386,05 € / 2,605 € = **148**.
  - Churn mensuel équivalent = 1 − (1 − 3,369 %)^(52/12) = **13,8 %**.
  - LTV simple = ARPU × marge / churn = 0,1036 / 0,03369 = **3,07 €**.
  - LTV : CAC = 3,07 / 2,61 = **1,18**.
  - Délai de remboursement du CAC (payback) = 2,61 / 0,1036 = **25 semaines**.
- **Lignes 35 à 37 : vérification avec la réalité.** La LTV mesurée par canal dans le notebook est de 2,03 € pour un utilisateur paid social, soit un **LTV:CAC de 0,78**.
  - Le modèle simple surestime la valeur d'un utilisateur payant, parce qu'il donne le même churn à tout le monde, alors que les utilisateurs paid social partent beaucoup plus vite.
  - C'est une limite importante à connaître.

### Model
Le calcul semaine par semaine du scénario choisi. La ligne 5 est la semaine 0 (22 juin 2026, dernière semaine observée) ; les lignes 6 à 57 sont les semaines 1 à 52.

| Colonne | Formule (ligne 6 = semaine 1) | Exemple Base, semaine 1 |
|---|---|---|
| C Start users | = utilisateurs de fin de la semaine précédente (I5) | 7 457 |
| D Churned | = C6 × churn | 7 457 × 3,369 % = 251 |
| E Retained | = C6 − D6 | 7 206 |
| M Paid spend | = budget semaine 1 × (1 + croissance)^(semaine − 1) | 386,05 € |
| F Paid new users | = M6 / CAC | 386,05 / 2,605 = 148 |
| G Free new users | = gratuits semaine 1 × (1 + croissance)^(semaine − 1) | 223 |
| H Referrals | = E6 × taux de parrainage | 0 |
| **I End users (WAU)** | = E + F + G + H | **7 577** |
| J WoW growth | = I6 / C6 − 1 | +1,6 % |
| K Revenue | = I6 × ARPU | 7 577 × 0,1036 = 785 € |
| L Gross profit | = K × marge | 785 € |
| N Cumul. GP − spend | cumul de (L − M) | 399 € |
| O Listening hours | = K / 0,009 € / 8,57 pubs par heure | 10 175 h |

La ligne 58 contient les totaux, O59 la moyenne des heures d'écoute, I60 le nombre d'utilisateurs actifs de la dernière semaine.

### Scenarios
- **Une ligne par scénario.**
  - Colonnes B à M : les paramètres, en vert car copiés d'Inputs.
  - Colonnes N (W0) à BN (W52) : les utilisateurs de chaque semaine.
  - Colonnes BO à BU : les résultats.
- **Formule d'une semaine**, par exemple O5 :
  `=N5*(1-$H5)*(1+$I5)+IF($E5>0;$C5*(1+$D5)^(O$4-1)/$E5;0)+$F5*(1+$G5)^(O$4-1)`
  - En clair : (actifs de la semaine d'avant) × (1 − churn) × (1 + parrainage) + nouveaux payants + nouveaux gratuits. C'est l'équation du professeur.
- **Budget pub sur 12 mois (BQ)** = somme d'une suite géométrique : 386,05 × ((1,00648)^52 − 1) / 0,00648 = **23 781 €**.
- **Lignes 12 à 18 : un seul levier change à la fois** (slide 98 du cours). Exemple : « churn −20 % » donne 15 589 utilisateurs en semaine 52, contre 13 490 en Base.
- **Lignes 22 à 29 : les 8 lignes « what-if » du notebook (section 9.4)** rejouées dans le modèle simple.

### Calibration
- **Pourquoi cet onglet :** le churn d'Echo n'est pas constant. Les nouveaux partent beaucoup dans les premières semaines, les anciens restent. Le modèle simple n'a qu'un seul churn. On choisit donc le **« churn équivalent »** : la valeur qui fait passer le modèle simple au plus près de la trajectoire du modèle par cohortes du notebook.
- **Méthode des moindres carrés :** on minimise SSE = Σ (WAU simple − WAU notebook)² sur les 52 semaines. Dans Excel, c'est `SUMXMY2(trajectoire simple; trajectoire notebook)` (colonne BC).
- **Pour chaque scénario (blocs aux lignes 26, 57 et 88) :**
  - ligne 2 du bloc : trajectoire du notebook (importée) ;
  - ligne 3 : trajectoire du modèle simple (en vert, copiée de Scenarios) ;
  - ligne 4 : écart en %.
  - Ensuite viennent l'erreur moyenne en % (MAPE, colonne BE) et l'écart final (colonne BF).
- **Le balayage (21 lignes)** recalcule la trajectoire avec un churn décalé de −0,5 à +0,5 point. La ligne « Churn with the lowest SSE » vérifie que le churn d'Inputs est bien le minimum : le texte « OK… » s'affiche.
- **Résultats :**

| Scénario | Churn équivalent | MAPE (erreur moyenne) | Écart en semaine 52 |
|---|---|---|---|
| Base | 3,369 % par semaine | 1,2 % | −1,7 % |
| Pessimiste | 5,073 % par semaine | 3,9 % | −6,4 % |
| Optimiste | 3,248 % par semaine | 1,5 % | −2,2 % |

- **Pour refaire le calcul dans Excel :** Données > Solveur (*Data > Solver*). Objectif : la cellule SSE du scénario, à minimiser, en modifiant le churn de ce scénario dans Inputs ligne 13.

### Backtest
Le modèle simple **seul**, sans l'aide du modèle par cohortes :
- Le churn est calé sur les utilisateurs actifs réels de juillet à décembre 2025 (C4 = 4,34 % par semaine).
- Le modèle prédit ensuite janvier à juin 2026 :
  - A : inscriptions prévues avec la tendance 2025 ;
  - B : inscriptions réelles.
- **Erreur (MAPE) :** 5,5 % (A) et 4,3 % (B) sur les utilisateurs actifs, contre 0,8 % et 2,1 % pour le modèle par cohortes.
- **Raison :** un seul churn ne voit pas que la base « vieillit » (plus d'anciens, qui partent moins).

### Checks
- Chaque résultat est comparé :
  - au notebook actuel (colonne C) ;
  - aux chiffres de votre consigne, issus de la version originale du notebook (colonne F).
- **Couleurs :** vert si l'écart est d'au plus 0,5 %, orange jusqu'à 5 %, rouge au-delà.
- Le rouge n'est pas un bug ici : il mesure ce que la version simple perd. Les écarts du backtest sont en points de pourcentage.

## 5. Glossaire

| Terme | Définition simple |
|---|---|
| WAU (Weekly Active Users) | Nombre d'utilisateurs qui ont ouvert l'app au moins une fois dans la semaine. |
| Churn | Part des utilisateurs actifs qu'on perd chaque semaine. Un churn de 3,4 % veut dire que sur 100 actifs, environ 3,4 ne reviennent pas la semaine suivante. |
| Rétention | L'inverse du churn : la part de ceux qui restent. |
| Churn équivalent | Churn constant qui imite au mieux un churn réel qui, lui, varie. |
| CAC (coût d'acquisition) | Dépense pub ÷ nouveaux utilisateurs payants. Ici 2,61 €. |
| ARPU | Revenu moyen par utilisateur actif, ici par semaine : 0,10 €. |
| LTV (lifetime value) | Revenu total qu'un utilisateur rapporte avant de partir. Formule simple : ARPU × marge ÷ churn. |
| LTV:CAC | Ce que rapporte un utilisateur ÷ ce qu'il coûte. En dessous de 1, on perd de l'argent ; le cours vise 3 ou plus. |
| Payback | Nombre de semaines pour rembourser le CAC. |
| Parrainage / referral | Nouveaux utilisateurs amenés par les utilisateurs existants. Echo n'a pas de données sur ce point, donc 0. |
| Cohorte | Groupe d'utilisateurs inscrits la même semaine. |
| Segment | Sous-groupe de même canal d'acquisition et même niveau d'activation (utilisé en version 2). |
| Activation | Un utilisateur est « activé » s'il suit au moins 3 podcasts en première semaine (le « aha moment » du notebook). |
| Facteur de calibration | Correction appliquée en version 2 pour que le point de départ colle aux vraies données. La version 1 n'en a pas besoin : elle part directement du vrai WAU. |
| Moindres carrés | Méthode qui choisit le paramètre minimisant la somme des écarts au carré. |
| SSE | Somme des carrés des écarts (*sum of squared errors*). |
| MAPE | Erreur moyenne en % (*mean absolute percentage error*). 2 % veut dire « en moyenne, on se trompe de 2 % ». |
| Backtest | On fait comme si on était dans le passé, on prédit une période déjà connue et on compare à la réalité. |
| North Star Metric | La métrique principale choisie dans le notebook : les heures d'écoute par semaine. |
| Scénario | Jeu d'hypothèses cohérent (Base, Pessimiste, Optimiste, Custom). |

## 6. Lien avec le notebook (section 9) et avec le simulateur du professeur

| Simulateur du professeur | Version 1 |
|---|---|
| Onglets Guide, Inputs, Model, Scenarios (+ Exercise) | Guide, Inputs, Model, Scenarios, plus Calibration, Backtest, Checks |
| `End = Start × (1 − churn) × (1 + referral) + spend / CAC` | Même équation, plus un terme « inscriptions gratuites » (celui de l'onglet *Exercise* du professeur), car 60 % des nouveaux utilisateurs d'Echo ne coûtent rien. |
| Pas mensuel, 12 mois | **Pas hebdomadaire, 52 semaines**, parce que le notebook mesure et prévoit des utilisateurs actifs *par semaine*. Un pas mensuel mélangerait des WAU avec un churn mensuel. |
| Colonnes Base / Optimistic / Pessimistic / Custom + SELECTED (INDEX/MATCH) | Identique, dans l'ordre Base / Pessimistic / Optimistic / Custom. |
| Unit economics (LTV, LTV:CAC, payback) | Identique, plus la LTV mesurée dans les données pour comparer. |
| Tests « un levier +20 % » | Identique, plus spend × 1,5, spend = 0, parrainage. |

**Traduction des hypothèses du notebook dans le modèle simple :**

| Hypothèse du notebook (9.3) | Dans le modèle simple | Ce qui est perdu |
|---|---|---|
| Nouveaux utilisateurs : rétention × ratio mesuré (0,72 après 15 semaines) | Absorbé dans le churn équivalent pessimiste (5,07 % au lieu de 3,37 %). | L'effet ne frappe pas que les nouveaux, et il ne dépend plus de l'âge. |
| Utilisateurs existants : −2,2 % d'activité par semaine pendant 26 semaines, puis arrêt | Absorbé lui aussi dans le churn pessimiste, mais **appliqué pendant les 52 semaines**. | La reprise après 26 semaines. D'où l'écart de −6,4 % en fin de période. |
| +6 % de revenu par semaine active | ARPU pessimiste = 0,1036 × 1,060 = 0,1098 €. | Rien. |
| +7 % de pubs par heure d'écoute | Pubs par heure pessimistes = 8,57 × 1,07. | Rien (sert seulement à calculer les heures d'écoute). |
| Activation +10 points (nouveaux utilisateurs) | Absorbé dans le churn équivalent optimiste (3,25 %). | L'effet n'est plus limité aux nouveaux et ne grandit plus avec le temps. |
| 6 segments, cohortes hebdomadaires, facteur de calibration | Un seul groupe d'utilisateurs, départ au vrai WAU. | Les différences entre canaux : un utilisateur paid social est traité comme les autres. |

## 7. Hypothèses, limites, et que répondre à « pourquoi ce modèle est-il valide ? »

**Hypothèses :**
- Churn constant ;
- parrainage nul ;
- marge 100 % (pas de données de coûts) ;
- budget pub et inscriptions gratuites qui croissent au même rythme que leur tendance 2025-2026 ;
- 52 semaines complètes : le notebook ne compte que 5 jours en semaine 1, d'où un budget pub supérieur de 2 %.

**Limites :**
- Le churn équivalent est **calé sur le modèle par cohortes** : ce n'est pas une mesure indépendante.
- Il colle mal au scénario pessimiste (erreur moyenne 3,9 %).
- Il **surestime la valeur de la pub payante** (dépense × 1,5 : +4 % d'utilisateurs de trop).
- Calé seulement sur l'historique, il sous-estime juin 2027 d'environ 18 %.
- Sa LTV simple (3,07 €) contredit la LTV mesurée d'un utilisateur payant (2,03 €).

**Réponse courte à l'oral :**
« Ce modèle n'est pas fait pour prédire seul. C'est un **résumé** du modèle par cohortes, dans le format du cours. Nous avons choisi le churn par moindres carrés pour qu'il reproduise la trajectoire du modèle détaillé, avec 1,2 % d'erreur moyenne en Base. Le modèle détaillé, lui, a été validé par backtest : 1 à 2 % d'erreur sur janvier-juin 2026. Le message tient dans les deux versions : le push agressif revient à faire passer le churn hebdomadaire de 3,4 % à 5,1 %, et le revenu à 12 mois baisse d'environ 16 %. Un taux de churn mesuré directement sur les données, 3,9 % par semaine sur le premier semestre 2026, est du même ordre que nos 3,4 %. »

## 8. Dépannage

| Problème | Solution |
|---|---|
| Des cellules affichent `####` | La colonne est trop étroite. Double-cliquez sur la bordure droite de la lettre de colonne. |
| Je change une valeur mais rien ne bouge | Le calcul est en mode manuel. Excel : Formules > Options de calcul > **Automatique** (ou touche F9). LibreOffice : Données > Calculer > Recalcul automatique. |
| Pas de flèche dans Inputs!C4 | Cliquez *dans* la cellule C4 : la flèche apparaît à droite. Si elle a disparu : Données > Validation des données > Autoriser : **Liste**, Source : `Base,Pessimistic,Optimistic,Custom`. |
| `#NOM?` / `#NAME?` | Vous avez sans doute tapé une formule en anglais dans un Excel en français ou l'inverse. Annulez avec Ctrl+Z. |
| `#REF!` | Une ligne ou une colonne a été supprimée. Fermez sans enregistrer, ou reconstruisez le fichier (voir le README principal). |
| LibreOffice affiche des 0 partout | Outils > Options > LibreOffice Calc > Formule > « Recalcul au chargement du fichier » : **Toujours**, puis rouvrez. Ou appuyez sur Ctrl+Maj+F9. |
| Les dates s'affichent comme des nombres (46 202) | Format de cellule > Date. |
| Virgule ou point décimal | Excel affiche selon la langue de votre ordinateur. Les valeurs sont les mêmes. |
| Le graphique ne bouge pas | Il suit la colonne I de Model : vérifiez que le recalcul est automatique. |
| J'ai tout cassé | Le fichier se reconstruit à l'identique avec `python build_workbooks.py` (voir le README principal). |

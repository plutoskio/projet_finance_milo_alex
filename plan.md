# Plan d'execution du projet finance - Michelin

## Objectif

Produire de facon autonome les deux livrables demandes dans `consigne.md` :

- un tableur Excel complet de valorisation de Michelin ;
- une presentation PowerPoint de 12 slides en francais.

L'entreprise analysee est Michelin, symbole boursier ML sur Euronext Paris.

## Etape 1 - Cadrage et structure du dossier

1. Lire la consigne et verifier les livrables attendus.
2. Creer une organisation simple du dossier :
   - `sources/` pour le journal des sources ;
   - `excel_model/` pour le modele de valorisation ;
   - `slides/` pour la presentation ;
   - `exports/` pour les elements intermediaires et controles.
3. Fixer le plan de slides impose par la consigne :
   - 1 slide d'executive summary ;
   - 3 slides d'analyse de marche ;
   - 3 slides d'analyse de l'entreprise ;
   - 5 slides de valorisation.

## Etape 2 - Collecte des informations

1. Collecter les sources primaires Michelin :
   - resultats annuels 2025 ;
   - guide des resultats 2025 ;
   - document d'enregistrement universel 2025 ;
   - pages officielles de gouvernance, actionnariat et donnees boursieres.
2. Collecter les donnees de marche :
   - taille du marche mondial du pneumatique ;
   - segmentation par type de produit et canal ;
   - croissance 2025 par segment ;
   - moteurs de croissance et risques.
3. Collecter les donnees de comparables cotes :
   - Bridgestone ;
   - Goodyear ;
   - Continental ;
   - Pirelli ;
   - Yokohama Rubber ;
   - Hankook Tire.
4. Collecter les inputs de valorisation :
   - cours de bourse Michelin ;
   - nombre d'actions ;
   - dette nette ;
   - taux sans risque ;
   - prime de risque actions ;
   - beta ;
   - cout de la dette et taux d'impot.

## Etape 3 - Analyse historique Michelin

1. Construire un historique 2020-2025 avec :
   - chiffre d'affaires ;
   - croissance ;
   - EBITDA ;
   - EBIT ;
   - resultat net ;
   - free cash-flow ;
   - capex ;
   - dette nette.
2. Calculer les ratios :
   - croissance du chiffre d'affaires ;
   - marge EBITDA ;
   - marge EBIT ;
   - marge nette ;
   - capex / chiffre d'affaires ;
   - free cash-flow / chiffre d'affaires ;
   - dette nette / EBITDA.
3. Rediger l'interpretation :
   - resilience des marges ;
   - pression sur les volumes 2025 ;
   - effet mix favorable ;
   - generation de cash-flow elevee.

## Etape 4 - Projections financieres

1. Projeter le P&L sur 2026-2030.
2. Utiliser des hypotheses explicites :
   - reprise progressive des volumes ;
   - croissance moderee du marche ;
   - effet mix premium ;
   - stabilisation puis amelioration des marges ;
   - poursuite de la discipline sur les capex et le cash-flow.
3. Construire les flux FCFF necessaires au DCF :
   - EBIT ;
   - impot theorique ;
   - NOPAT ;
   - depreciation et amortissement ;
   - capex ;
   - variation de BFR ;
   - FCFF.

## Etape 5 - Valorisation par comparables

1. Justifier l'echantillon de pairs : fabricants mondiaux de pneumatiques et equipementiers ayant une exposition similaire aux cycles auto, remplacement, B2B et specialites.
2. Calculer les multiples :
   - EV / Chiffre d'affaires ;
   - EV / EBITDA ;
   - EV / EBIT.
3. Appliquer une approche prudente :
   - utiliser la mediane comme point central ;
   - afficher un bas de fourchette et un haut de fourchette ;
   - exclure ou commenter les distorsions liees aux pairs tres endettes ou diversifies.

## Etape 6 - Valorisation DCF

1. Calculer le WACC :
   - taux sans risque francais ;
   - prime de risque Eurozone ;
   - beta Michelin ;
   - cout de la dette apres impots ;
   - structure de capital de marche.
2. Actualiser les FCFF 2026-2030.
3. Calculer la valeur terminale avec une croissance perpetuelle prudente.
4. Passer de l'enterprise value a l'equity value :
   - enterprise value ;
   - moins dette nette ;
   - equity value ;
   - valeur par action.
5. Construire une sensibilite WACC / croissance terminale.

## Etape 7 - Recommandation d'investissement

1. Comparer la valorisation issue des comparables et du DCF au cours de bourse actuel.
2. Calculer le potentiel de hausse ou de baisse.
3. Comparer la rentabilite attendue au cout des fonds propres.
4. Conclure avec une recommandation claire : Acheter, Conserver ou Vendre.
5. Presenter les principaux catalyseurs et risques.

## Etape 8 - Production des livrables

1. Generer le modele Excel avec les onglets suivants :
   - `Sources` ;
   - `Historical` ;
   - `Forecast` ;
   - `Comps` ;
   - `WACC` ;
   - `DCF` ;
   - `Sensitivity` ;
   - `Football_Field`.
2. Generer une presentation PowerPoint de 12 slides :
   - slide 1 : Executive summary ;
   - slide 2 : Taille du marche et segmentation ;
   - slide 3 : Concurrence et barrieres a l'entree ;
   - slide 4 : Croissance, drivers et risques ;
   - slide 5 : Activite, actionnariat et management ;
   - slide 6 : P&L historique ;
   - slide 7 : Projection du P&L ;
   - slide 8 : Comparables retenus ;
   - slide 9 : Resultats des comparables ;
   - slide 10 : WACC et hypotheses DCF ;
   - slide 11 : Resultats du DCF ;
   - slide 12 : Football field et recommandation.

## Etape 9 - Controle qualite

1. Verifier que la presentation contient exactement 12 slides.
2. Verifier la coherence entre Excel et PowerPoint :
   - chiffre d'affaires ;
   - marges ;
   - WACC ;
   - dette nette ;
   - cours de bourse ;
   - valeurs par action.
3. Verifier que les formules Excel fonctionnent.
4. Verifier que chaque slide contient les messages essentiels attendus par la consigne.
5. Ajouter les sources et avertissements de methode dans le modele et dans les notes ou bas de slides.


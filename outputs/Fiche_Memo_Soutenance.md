# FICHE MÉMO — SOUTENANCE
## The Role of Weather in French Day-Ahead Electricity Price Forecasting
**Leo Cambreleng & Lyam Oumedjeber · EDHEC MSc DAAI · Prof. Milos Vulanovic · Juin 2026**

---

## LA PHRASE THESIS EN 1 LIGNE
> *"La météo est **redondante** en régime stable (déjà dans load forecast + TTF) mais **significative** en crise 2022 (canal gaz décorrélé + contraintes nucléaires)."*

---

## RÉSULTATS CLÉS — 2 RÉGIMES

| | **Stable 2024–25** *(n=8 640)* | **Crise 2022** *(n=8 760)* |
|---|---|---|
| **Naïf A — MAE** | 33,09 EUR/MWh | 72,74 EUR/MWh |
| **RF sans météo B — MAE** | 16,95 · R²=0,775 | 67,28 · R²=0,464 |
| **RF avec météo C — MAE** | 16,94 · R²=0,776 | 66,55 · R²=0,475 |
| **XGBoost D — MAE** | 19,14 · R²=0,659 | 76,32 · R²=0,282 |
| **DM test C vs B** | p=0,572 · stat=−0,565 · **n.s.** | p<0,001 · stat=**−13,27 · *** |
| **ML vs naïf** | −49% MAE · p<0,001 | −8,5% MAE · p<0,001 |

---

## TESTS DIEBOLD-MARIANO (régime stable)

| Comparaison | DM stat | p-value | Verdict |
|---|---|---|---|
| C vs B — météo ajoute de la valeur ? | −0,565 | 0,572 | **n.s.** |
| C vs A — RF-météo bat naïf ? | −53,53 | <0,001 | *** |
| B vs A — RF-sans-météo bat naïf ? | −53,44 | <0,001 | *** |
| D vs C — XGBoost vs RF-météo ? | +10,88 | <0,001 | *** (XGB pire) |

---

## FEATURE IMPORTANCE (RF avec météo, MDI)

| Feature | Importance |
|---|---|
| price_lag24h | **20,7%** |
| price_roll24h_mean | 14,3% |
| price_lag48h | 12,4% |
| price_roll168h_mean | 11,9% |
| price_lag168h | 10,4% |
| **Lags + rolling means TOTAL** | **~70%** |
| TTF gaz | 8,3% |
| ARA coal | 6,0% |
| nuclear_avail_ratio | ~1,9% |
| **TOUTE la météo combinée** | **~2,2%** |

---

## BACKTEST TRADING (scénario central 0,30 EUR/MWh, 1 MW)

| Modèle | Net P&L (EUR/MW) | Sharpe | MaxDD | Calmar |
|---|---|---|---|---|
| A — Naïf | 85 482 | 10,53 | 2 367 | 36,6 |
| B — RF sans météo | 136 884 | 19,44 | 433 | 320,6 |
| **C — RF avec météo** | **136 711** | **19,46** | **422** | **328,3** |
| D — XGBoost | 142 012 | 18,78 | 1 351 | 106,6 |
| Long-only | ≈ 0 | −0,01 | 4 521 | — |

**RF drawdown 5,6× plus petit que naïf** · Sharpe positif dans **tous les 12 mois**

---

## SETUP DONNÉES & MODÈLES

- **Période** : Jan 2018 – Avr 2025 · ~64 000 obs horaires · 4 sources publiques
- **Sources** : ENTSO-E (prix/charge/prod/flux) · ERA5 ECMWF (météo) · TTF gas · ARA coal
- **Features** : 35 total (27 sans météo pour B) — calendrier, lags prix, fondamentaux, flux, météo, dérivées
- **RF** : 500 arbres · profondeur 10 · min-leaf 5 · règle sqrt · random_state=42
- **XGBoost** : 500 est. · η=0,05 · profondeur 6 · subsample=0,8 · colsample=0,8
- **Test stable** : mai 2024 – avr 2025 | **Test crise** : train 2018–2021, test 2022 entier
- **DM** : HLN-corrected, MAE-based loss differential

---

## FRANCE : SPÉCIFICITÉS À CONNAÎTRE

- ~63 GW nucléaire installé ≈ **~70% de la production**
- Thermosensibilité **~2,4 GW/°C** (la plus haute d'Europe, source RTE 2023)
- **HDD seuil 17°C** (méthode RTE) · **CDD seuil 22°C**
- Disponibilité nucléaire 2022 : **~40%** (corrosion sous contrainte, chocs historiques)
- Disponibilité nucléaire 2024–25 : **~65%** (normalisée)
- Dead-band trading : **δ = 2 EUR/MWh** ≈ 20e percentile du mouvement absolu J vs J-1

---

## 3 MÉCANISMES DE REDONDANCE (régime stable)

1. **Load forecast ≡ température** — RTE publie une prévision de charge construite sur des prévisions météo → inclure cette prévision = conditionner déjà sur la météo
2. **TTF gas ≡ météo européenne** — le gaz agrège la demande de chauffage pan-européenne → TTF capture la météo comme proxy commodité
3. **Domination nucléaire** — le nucléaire explique une large part de la variance des prix indépendamment de la météo → peu de résidu pour la météo

---

## POURQUOI LA CRISE 2022 CASSE LA REDONDANCE

1. **TTF décorrèle de la température locale** — choc géopolitique russe → TTF monte sans lien avec la météo française → la variable météo doit maintenant l'encoder directement
2. **Amplification nucléaire** — à 40% de disponibilité, chaque degré froid de plus → urgence sur les centrales gaz/fioul → pas de marge d'absorption
3. **Signal/bruit monte** — les prix à 200–1 000 EUR/MWh créent des différentiels d'erreur B vs C mesurables (0,73 EUR/MWh absolus = detectables par DM)

---

## PIÈGES À ÉVITER (les 5 confusions mortelles)

| Ne pas dire | Dire à la place |
|---|---|
| "La météo est inutile" | "Redondante en régime calme, significative en crise" |
| "Notre Sharpe de 19,5 est exceptionnel" | "C'est 1 MW sans levier, non comparable aux fonds actions — regardez le Calmar et le drawdown" |
| "ERA5 est une prévision météo" | "ERA5 est une réanalyse — c'est la météo parfaite réalisée, une borne supérieure" |
| "XGBoost est inférieur à RF" | "Avec nos hyperparamètres par défaut — un tuning Bayésien pourrait inverser" |
| "Le prix de demain = prix d'aujourd'hui" | "Le lag-24h est la feature la plus importante (20,7%) mais 29 autres features contribuent aux 79,3% restants" |

---

*Imprime cette page, garde-la sur la table. Tout le reste est dans ta tête.*

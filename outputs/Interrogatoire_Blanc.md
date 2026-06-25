# Interrogatoire Blanc — Soutenance Master Thesis
## The Role of Weather in French Day-Ahead Electricity Price Forecasting
**Leo Cambreleng & Lyam Oumedjeber — EDHEC MSc DAAI — Juin 2026**

---

## PARTIE A — 20 Questions avec réponses complètes

> Entraîne-toi à répondre à voix haute en moins de 2 minutes chacune. La réponse type : 1 phrase directe → justification → chiffre clé.

---

### Q1 — Votre résultat principal dit que la météo "ne sert pas en régime stable". N'est-ce pas une conclusion négative un peu décevante ?

**Réponse :**
Non — c'est une conclusion précise, pas négative. Elle dit *quand* et *pourquoi* la météo est redondante. Le vrai résultat a deux parties indissociables : redondance en régime stable (p = 0,572), mais haute significativité en crise 2022 (p < 0,001, DM = −13,27). Une thèse qui dit "parfois oui, parfois non, et voici le mécanisme" est scientifiquement plus riche qu'une thèse qui dit simplement "oui". Le résultat a une implication opérationnelle directe : il faut un détecteur de régime, pas une inclusion systématique ou une exclusion systématique.

---

### Q2 — Pourquoi avoir choisi le Random Forest comme modèle principal plutôt qu'un LSTM ou un Transformer ?

**Réponse :**
Trois raisons adaptées à notre problème. Premièrement, RF capture naturellement les non-linéarités à seuil — notamment le seuil HDD à 17°C — via ses splits. Deuxièmement, RF est robuste aux valeurs extrêmes (prix > 1 000 EUR/MWh en 2022) car il moyenne sur 500 arbres bootstrappés, ce qui dilue l'influence des outliers. Troisièmement, la mesure MDI nous donne une importance par feature qui est essentielle à notre question de recherche : combien vaut la météo ? Un LSTM ou Transformer ne fournit pas cette interprétabilité native. Enfin, avec 64 000 observations, le ratio signal/paramètres favorise RF. Les architectures profondes sont mentionnées comme piste future (ch. 6.5).

---

### Q3 — Le test de Diebold-Mariano avec correction HLN — pourquoi cette correction spécifiquement ?

**Réponse :**
La correction Harvey-Leybourne-Newbold (1997) est un ajustement de taille d'échantillon fini appliqué à la statistique DM originale. La statistique DM est asymptotiquement normale, mais sur des horizons courts l'approximation normale surestime la significativité — la distribution réelle a des queues plus épaisses. HLN remplace la loi normale par une loi de Student dont les degrés de liberté dépendent de n. Sur n = 8 640 observations, la différence est faible mais réelle, et la correction est recommandée par Diebold & Mariano eux-mêmes comme pratique standard. Elle rend notre test plus conservateur, ce qui renforce la conclusion de non-significativité en régime stable.

---

### Q4 — L'ERA5 est une réanalyse, pas une prévision météo. N'introduisez-vous pas un biais look-ahead ?

**Réponse :**
Point méthodologique important. ERA5 utilise des observations météo réelles et les assimile pour produire la "meilleure estimation" du temps passé. Ce n'est pas une prévision : en production réelle, on n'aurait que des prévisions NWP (ECMWF, Météo-France) avec une erreur de 10-30%. Notre design teste donc la valeur *maximale théorique* de la météo — une borne supérieure. Si même la météo parfaite est non-significative en régime stable (p = 0,572), une vraie prévision météo imparfaite ne l'aurait été qu'encore moins. Cela renforce notre conclusion en régime calme. En revanche, pour la crise 2022 où le résultat est fortement significatif, il faut nuancer : la vraie prévision météo aurait donné un gain plus faible qu'ERA5, mais très probablement encore significatif (DM = −13,27 laisse beaucoup de marge).

---

### Q5 — Votre Sharpe annualisé de 19,5 est irréaliste pour un vrai fonds. Pourquoi le présentez-vous ?

**Réponse :**
Le Sharpe n'est pas présenté comme une performance absolue comparable à un fonds d'actions — nous le disons explicitement (section 5.3, "Note on Sharpe interpretation"). Ce chiffre est une mesure *relative entre modèles, à taille et levier identiques* : 1 MW, sans levier, sans déduction du taux sans risque (convention marché énergie). Les marchés électriques sont mean-reverting à long terme (le long-only gagne ~0), donc la comparaison fonds-actions est structurellement invalide. La métrique pertinente est le **Calmar Ratio** (rendement annualisé / MDD) : RF à 328 vs naïf à 37 — c'est la robustesse que l'on veut montrer. Et le drawdown de 422 EUR/MW vs 2 367 pour le naïf : voilà la vraie mesure du risque. Le Sharpe positif dans **tous les 12 mois** montre la consistance.

---

### Q6 — Vous n'avez pas inclus les prix EUA (quotas carbone). N'est-ce pas un facteur déterminant ?

**Réponse :**
Nous reconnaissons cette limite en section 6.3. L'argument central : l'EUA influence le coût marginal du gaz et du charbon, qui sont eux-mêmes inclus (TTF, ARA coal). Le spread gaz-charbon en EUR/MWh capture déjà une grande partie du signal carbone car les trois séries ont évolué ensemble en 2021-2022. Exclure l'EUA explicitement n'efface pas sa contribution — elle est absorbée par le spread. De plus, le test d'ablation C vs B change seulement les features météo, pas les features fondamentaux : même avec EUA ajouté, la comparaison météo/sans-météo resterait valide. Ajouter EUA est une extension naturelle de faible coût, mais nous ne pensons pas qu'elle changerait notre conclusion principale.

---

### Q7 — XGBoost est généralement meilleur que Random Forest dans les benchmarks. Votre résultat est-il fiable ?

**Réponse :**
Oui, avec une nuance importante que nous présentons explicitement. L'infériorité d'XGBoost ici reflète nos hyperparamètres par défaut, non une infériorité intrinsèque d'XGBoost. Le mécanisme : le boosting séquentiel sur-pondère les résidus des heures de prix extrêmes (> 500 EUR/MWh en 2022 qui font partie du train) lors des dernières itérations, ce qui crée de l'overfitting sur ces outliers. RF, en moyennant sur 500 bootstrap, dilue cet effet. En 2022, l'écart se creuse encore : XGBoost MAE 76,32 vs RF 66,55 (−14,6% de dégradation supplémentaire). Un tuning Bayésien sur XGBoost (learning rate, λ, α, colsample) pourrait fermer ou inverser le gap. La comparaison B vs C (RF sans/avec météo) reste propre car même architecture.

---

### Q8 — La "thermosensibilité de 2,4 GW/°C" — d'où vient ce chiffre et est-il fiable ?

**Réponse :**
Ce chiffre est publié par RTE dans le Bilan Électrique 2023 (notre référence rte2023). C'est une estimation de la sensibilité de la demande totale française à une variation de 1°C de la température, mesurée sur les mois d'hiver. Il est calculé empiriquement sur la relation historique load forecast / température. La fiabilité est bonne car : (1) RTE est l'opérateur du réseau, il a accès à toutes les données de consommation, (2) c'est une méthodologie établie utilisée dans leurs rapports annuels depuis des années. C'est la thermosensibilité la plus élevée d'Europe car environ 35% du parc de chauffage français est électrique, contre ~8% en Allemagne.

---

### Q9 — Votre test d'ablation C vs B n'est valide que si les modèles B et C sont par ailleurs identiques. Comment vous l'assurez-vous ?

**Réponse :**
Par construction : B et C utilisent exactement les mêmes hyperparamètres (500 arbres, profondeur 10, min-leaf 5, règle sqrt, random_state=42), le même train set (janvier 2018 — avril 2024), le même test set (mai 2024 — avril 2025), et le même critère d'évaluation (MAE-based DM avec HLN). La seule différence est la matrice de features : 27 pour B, 35 pour C (B + 8 variables météo). Le DM test compare directement les erreurs heure par heure sur les mêmes 8 640 observations. C'est un design d'ablation propre — exactement l'analogie d'une expérience contrôlée où l'on change une seule variable.

---

### Q10 — Pourquoi utiliser le lag-168h pour votre benchmark naïf plutôt que le lag-24h plus courant ?

**Réponse :**
Le lag-168h est plus exigeant que le lag-24h, donc plus honnête. Les prix électriques ont une autocorrélation hebdomadaire très forte — le prix à 8h lundi suit mieux le prix à 8h du lundi précédent que le prix à 8h du dimanche précédent, car les patterns de demande sont hebdomadaires (pic lundi-vendredi, creux week-end). Le lag-168h capture cette saisonnalité hebdomadaire mieux que le lag-24h. Un modèle ML qui bat seulement le lag-24h a une barre basse. En battant le lag-168h avec une réduction de MAE de 49% (33,09 → 16,94), nous démontrons une valeur ajoutée plus robuste. C'est devenu la pratique standard en EPF (Weron 2014, Lago 2021).

---

### Q11 — Les prix négatifs représentent 3-8% des observations. Vos modèles les traitent-ils correctement ?

**Réponse :**
Partiellement. Notre métrique principale MAE est symétrique et ne fait pas de discrimination positive/négative — elle traite les prix négatifs comme n'importe quelle observation. Nous utilisons sMAPE plutôt que MAPE pour éviter la division par des valeurs proches de zéro. Cependant, nos modèles RF et XGBoost ciblent l'espérance conditionnelle : ils prédisent la moyenne, pas le régime négatif spécifique. En régime de prix négatifs, le mécanisme de formation est différent — les producteurs inflexibles (nucléaire, éolien) paient pour maintenir l'équilibre réseau. Ce régime spécifique mériterait un modèle dédié. C'est mentionné comme limite en section 6.3. Aucune correction ou filtrage spécial n'a été appliqué.

---

### Q12 — Votre dataset commence en 2018. Qu'est-ce qui justifie ce choix plutôt que 2015 ou 2010 ?

**Réponse :**
Deux raisons. Première : la disponibilité des données ERA5 via l'API Copernicus CDS à résolution horaire sur la France avec une qualité homogène commence pratiquement à être exploitable à partir de 2018 (couverture plus stable, variables supplémentaires). Deuxième : la structure du marché français a évolué — le parc solaire et l'éolien ont significativement augmenté, les règles EPEX ont changé. Partir de 2018 garantit une certaine homogénéité structurelle du marché. Commencer avant 2015 aurait inclus des données où la pénétration des renouvelables était beaucoup plus faible, créant une distribution des prix non-stationnaire qui aurait biaisé l'apprentissage. Le choix de 2018 est cohérent avec la littérature récente (Tschora et al. 2022, Lago et al. 2021).

---

### Q13 — Comment interpretez-vous le fait que le lag-24h ait 20,7% d'importance mais que votre modèle ne soit pas simplement "le prix d'hier" ?

**Réponse :**
Le lag-24h étant la feature la plus importante ne signifie pas que le modèle est naïf — cela signifie qu'il est bien calibré. Le lag-168h (benchmark naïf) a seulement 10,4% d'importance dans le MDI malgré sa forte saisonnalité, parce que le lag-24h capture déjà la majeure partie de l'autocorrélation journalière. La valeur ajoutée du RF vient de la **combinaison** de toutes les features : 70% pour les lags et rolling means, 8,3% TTF, 6% charbon, etc. Un modèle "simplement le lag-24h" serait le benchmark lag-24h, non inclus car moins standard que le lag-168h — mais nos RF battent les deux. L'importance MDI mesure la contribution marginale *étant donné les autres features*, pas la corrélation univariée.

---

### Q14 — En 2022, votre MAE passe de 16,94 à 66,55 EUR/MWh. Est-ce que votre modèle est vraiment utile dans ce contexte ?

**Réponse :**
Oui — la comparaison pertinente n'est pas la MAE absolue mais le gain *versus* le benchmark. En 2022, le benchmark naïf atteint MAE = 72,74 EUR/MWh. Notre modèle RF à 66,55 représente quand même une réduction de 8,5%. Le marché est intrinsèquement imprévisible en crise (chocs nucléaires, guerre, prix du gaz × 10) — aucun modèle de régression ne prédit des événements idiosyncratiques. Ce qui importe : (1) nos modèles restent les meilleurs disponibles avec cette architecture, (2) le R² passe de 0,160 (naïf) à 0,475 (RF), soit une amélioration substantielle malgré la crise. En trading, la dégradation de précision se retrouve dans le Sharpe 2022 plus faible — mais un modèle robuste reste une aide précieuse même imparfaite.

---

### Q15 — La stratégie de trading suppose des exécutions parfaites. En pratique, comment les coûts implicites réduisent-ils la performance ?

**Réponse :**
Nous modélisons les coûts *explicites* de transaction EPEX (0,10 à 0,60 EUR/MWh selon le scénario) mais pas les coûts *implicites* : market impact, slippage, partial fills en heures off-peak. Pour un book de 1 MW, le market impact est négligeable sur EPEX (liquidité élevée, marché de volume). Pour des positions plus importantes (50-100 MW), le cost compression serait réel. La bonne nouvelle : à 0,60 EUR/MWh (scénario pessimiste), le Sharpe reste 19,16 vs 19,46 en optimiste — soit une dégradation de <2%. La robustesse aux coûts est donc démontrée dans la fourchette réaliste. Les tableaux de sensibilité sont en section 5.4.2.

---

### Q16 — Avez-vous considéré d'autres marchés européens pour tester la généralisation de votre hypothèse de redondance ?

**Réponse :**
Pas dans cette étude — c'est délibéré et constitue notre principale extension future. L'hypothèse de redondance est spécifique à la structure française : nucléaire dominant (~70%), thermosensibilité électrique élevée, et TTF comme proxy météo efficace. En Allemagne, les renouvelables dépassent 60% de la production (Fraunhofer ISE 2025) et l'éolien+solaire génèrent directement un canal météo-offre beaucoup plus fort que le canal météo-demande. Notre hypothèse prédit que la météo serait plus significative sur EPEX Allemagne qu'en France, même en régime stable. Une étude comparative France-Allemagne avec le même protocole expérimental est la piste la plus directe pour tester la généralité de notre mécanisme.

---

### Q17 — La correction HLN s'applique à des erreurs à horizon h fixe. Vos 8 640 observations sont 8 640 heures mais pas indépendantes. Le test est-il valide ?

**Réponse :**
Le test DM est conçu pour des séries avec autocorrélation — c'est précisément son avantage sur une comparaison naïve de moyennes. La statistique DM est basée sur les différentiels de perte $d_t = |e_t^{(1)}| - |e_t^{(2)}|$ qui peuvent être autocorrélés. Diebold & Mariano (1995) et Harvey et al. (1997) proposent un estimateur de variance des $d_t$ qui est robuste à l'autocorrélation (via un noyau Bartlett de type Newey-West — notre référence newbold1993). Donc le test tient compte de la dépendance sérielle. Le seul point délicat est le choix du truncation lag pour l'estimateur de variance : nous utilisons la valeur par défaut recommandée pour les prévisions un pas en avant.

---

### Q18 — Vous affirmez que la méthode est reproductible. Que faudrait-il pour que quelqu'un reproduise exactement vos résultats ?

**Réponse :**
Quatre éléments : (1) **Données** — tout vient de sources publiques (ENTSO-E Transparency Platform, Copernicus CDS ERA5, TTF via marchés publics, ARA coal via Bloomberg/Quandl) avec les paramètres d'appel API précisés. (2) **Code** — disponible dans le repo GitHub du projet avec toutes les étapes de feature engineering. (3) **Graines aléatoires** — random_state=42 pour RF et XGBoost. (4) **Découpe temporelle** — train : janvier 2018–avril 2024 ; test stable : mai 2024–avril 2025 ; test crise : train 2018-2021, test 2022. La reproductibilité est une valeur de notre travail — c'est l'une des différences avec beaucoup de papiers EPF qui n'ouvrent pas leur code ou leurs données.

---

### Q19 — La métrique sMAPE : pourquoi pas simplement le RMSE qui est plus standard ?

**Réponse :**
Nous *reportons* le RMSE mais ne l'utilisons pas comme critère principal. RMSE pénalise quadratiquement les grandes erreurs — il est dominé par les heures de prix extrêmes qui représentent <1% des observations mais une fraction énorme de l'erreur. Un modèle qui prédit parfaitement toutes les heures normales mais échoue sur les spikes aurait un RMSE élevé mais un MAE acceptable. Le MAE est plus représentatif de la performance médiane. Le sMAPE est un indicateur de pourcentage robuste aux valeurs proches de zéro (MAPE pur diverge quand le prix réel → 0 ou négatif, ce qui arrive 3-8% des heures). Tous les cinq métriques sont reportés pour la transparence — le jury peut voir le tableau complet en Table 4.1.

---

### Q20 — Si vous recommencez cette thèse, que changeriez-vous en priorité ?

**Réponse :**
Trois choses. Premièrement, le **retraining walk-forward** : entraîner mensuellement sur une fenêtre expansive plutôt que batch unique — plus proche du production. Deuxièmement, l'**optimisation Bayésienne d'XGBoost** avec temporal cross-validation : la comparaison RF/XGBoost serait plus honnête avec XGBoost bien tuné. Troisièmement, inclure les **prix EUA** explicitement et tester sur la **période 2023** comme deuxième régime stable (différent de 2024-25 : encore des séquelles de la crise, nucléaire remontant). Ces extensions ne remettent pas en cause la conclusion principale — elles la renforceraient. Le résultat 2022 est robuste : DM = −13,27 est un effet massif qui survivrait à tout raffinement méthodologique raisonnable.

---

## PARTIE B — 5 Questions à ne pas traiter seul(e)

> Ces questions sont délicates car elles peuvent aller dans plusieurs directions. Concerte-toi avec l'autre avant de répondre — ou répartis-vous la réponse.

---

### B1 — "Vous concluez que la météo ne sert pas. Avez-vous testé d'autres spécifications météo — ARPEGE, COSMO, prévisions NWP ?"

**Pourquoi c'est piégeux :** Si tu dis "non, seulement ERA5", tu sembles avoir fait un choix non justifié. Si tu dis "oui" et que tu n'as pas testé, tu mens.

**Comment gérer à deux :** Reconnais clairement que c'est ERA5 uniquement (réanalyse = borne supérieure). Donne l'argument ERA5 = upper bound. Si le jury insiste, Leo peut prendre la direction technique (pourquoi ERA5 est la bonne baseline) pendant que Lyam reformule l'implication (si même la météo parfaite ne change pas le résultat, la météo imparfaite d'une NWP ne le changerait qu'encore moins en régime stable).

---

### B2 — "Votre dataset couvre une période de 7 ans. Comment gérez-vous la non-stationnarité ? Avez-vous testé la stabilité des coefficients dans le temps ?"

**Pourquoi c'est piégeux :** RF n'a pas de "coefficients" au sens linéaire. Parler de stationnarité pour un RF demande de passer par les importances ou les performances rolling.

**Comment gérer à deux :** Lyam peut répondre sur les Sharpe mensuels (positifs tous les 12 mois = stabilité relative) et Leo peut aborder la stationnarité des features (les lags de prix capturent des autocorrélations robustes au-delà du niveau des prix, qui lui change). Mentionner le rolling MAE par mois (Figure 4.4) comme proxy de stabilité.

---

### B3 — "Avez-vous envisagé des modèles non-supervisés (clustering des régimes) pour détecter automatiquement les crises ?"

**Pourquoi c'est piégeux :** C'est une extension naturelle que vous n'avez pas faite. Il faut ni prétendre l'avoir faite, ni rejeter l'idée.

**Comment gérer à deux :** Lyam reformule l'idée positivement : "C'est exactement ce que nous envisageons comme détecteur de régime dans notre recommandation finale (ch. 6.1). Un clustering HMM ou k-means sur la disponibilité nucléaire + volatilité TTF pourrait automatiser ce que nous faisons ici manuellement avec 2022 comme régime de test." Leo peut développer l'opérationnalisation : les signaux observables (nuclear_avail_ratio < 0,55, volatilité TTF).

---

### B4 — "La contribution marginal de la météo en crise est 0,73 EUR/MWh (67,28 → 66,55). Économiquement, est-ce significatif ?"

**Pourquoi c'est piégeux :** 0,73 EUR/MWh sur une MAE de 67 EUR/MWh, c'est 1,1% — en valeur relative, faible. Pourtant le DM est −13,27 avec p < 0,001. Le jury cherche à voir si vous distinguez significativité *statistique* vs *économique*.

**Comment gérer à deux :** Leo prend le côté statistique (DM = −13,27 est un effet massif *en rapport au bruit de mesure*). Lyam prend le côté économique : en période de crise à 100-200 EUR/MWh, un gain de 0,73 EUR/MWh moyen *masque* des moments où l'avantage est plus grand. Le MAE moyen ne capture pas les heures de pointe où la météo fait vraiment la différence. L'argument de robustesse : même si on doute de la signification économique absolue, le *changement de régime* (n.s. → ***) est le résultat, pas la magnitude en EUR.

---

### B5 — "Vous avez utilisé la même architecture RF pour B et C. Un modèle dédié météo (SARIMA-GARCH + météo) n'aurait-il pas été plus puissant ?"

**Pourquoi c'est piégeux :** Il y a des arguments valides dans les deux sens. Le jury teste si vous avez pesé l'alternative.

**Comment gérer à deux :** Leo défend le choix : l'ablation B vs C est valide *précisément parce que* l'architecture est identique — elle isole la contribution des features. Comparer C (RF) à un SARIMA-météo comparerait *architecture ET features* simultanément, rendant l'ablation impossible. Lyam peut mentionner que Lago et al. (2021) montrent que RF et modèles économétriques ont des performances comparables sur les marchés européens — RF n'est pas sous-performant en absolu.

---

## PARTIE C — 10 Phrases de Récupération

> Pour les moments de blanc, de surprise, ou d'attaque rhétorique. Prononce-les calmement, gagne du temps, et reviens à tes chiffres.

---

1. **Quand tu ne sais pas la réponse exacte :**
   > "C'est une très bonne question que nous avons effectivement discutée lors de notre analyse — permettez-moi d'y réfléchir une seconde pour vous donner la réponse précise plutôt qu'approximative."

2. **Quand tu confonds un chiffre :**
   > "Je veux vous donner le chiffre exact — dans notre thèse, la valeur reportée est [pause, reprends] — plutôt que risquer une approximation sur une statistique centrale."

3. **Quand tu ne comprends pas la question :**
   > "Pour être sûr de répondre à ce que vous demandez : est-ce que votre question porte principalement sur [reformulation A] ou sur [reformulation B] ?"

4. **Quand on t'attaque sur une limite :**
   > "Vous avez tout à fait raison d'identifier cette limite — nous la documentons nous-mêmes en section 6.3. La question est si elle invalide le résultat principal, et nous pensons que non, pour la raison suivante..."

5. **Quand tu dois rendre la main à l'autre :**
   > "Leo/Lyam a travaillé sur cet aspect spécifique de l'implémentation — je lui passe la parole pour la précision technique."

6. **Quand la question est très large :**
   > "C'est un point qui touche à plusieurs aspects de notre méthodologie. Pour répondre précisément : concernant [aspect 1], voici ce que nous avons fait ; concernant [aspect 2]..."

7. **Quand tu as fait une erreur dans la présentation :**
   > "Merci de le signaler — vous avez raison, la formulation sur la slide [X] n'est pas précise. La valeur correcte est [Y] comme reporté dans le tableau [Z]."

8. **Quand la question sort du périmètre de la thèse :**
   > "Cette question dépasse le périmètre de notre étude — nous nous sommes volontairement limités à [France / 2018-2025 / RF+XGBoost]. C'est d'ailleurs précisément pourquoi nous proposons [l'extension] comme piste future."

9. **Pour conclure n'importe quelle réponse en revenant au cœur :**
   > "En résumé : notre résultat central — la dépendance au régime de la valeur météo — reste intact. [Chiffre clé]. [Mécanisme en une phrase]."

10. **Pour les silences gênants après une question difficile :**
    > "C'est précisément le genre de question qui nous a amenés à pousser l'analyse vers la validation 2022 — laissez-moi vous expliquer comment nous y avons répondu empiriquement."

---

*Bon courage pour la soutenance — la thèse est solide, les chiffres parlent d'eux-mêmes.*

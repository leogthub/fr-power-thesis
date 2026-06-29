# SCRIPT DE SOUTENANCE — MOT POUR MOT
## The Role of Weather in French Day-Ahead Electricity Price Forecasting
**Leo Cambreleng & Lyam Oumedjeber · EDHEC MSc DAAI · Juin 2026**

> **FORMAT** : 20 min exposé + 10 min Q&A · 1 examinateur · Teams · Anglais  
> **LEO** parle slides 1–7 · **LYAM** parle slides 8–15 · Slide 16 : les deux  
> Timing indicatif entre crochets — total visé **~19 minutes**

---

## SLIDE 1 — TITRE `[LEO · ~30s]`

Good morning, Professor. Thank you for joining us.

My name is Leo Cambreleng, and alongside my co-author Lyam Oumedjeber, I am presenting our Master's thesis from the MSc in Data Analysis and AI at EDHEC Business School.

Our thesis is titled: *"The Role of Weather in French Day-Ahead Electricity Price Forecasting: A Random Forest Approach."*

We were supervised by Professor Milos Vulanovic. I will cover the first half of the presentation — context, data, and the stable-regime results — and Lyam will take over for the crisis analysis, the trading backtest, and our conclusions.

---

## SLIDE 2 — CONTEXT & MOTIVATION `[LEO · ~2min]`

Let me start with the question that motivated this work: why is forecasting French electricity prices so difficult — and why might weather matter?

First, electricity cannot be stored at scale. This means supply and demand must clear instantaneously every single hour. The result is a price series that is unlike any other financial asset: extreme volatility, strong intraday seasonality, and spikes that can reach several thousand euros per megawatt-hour during scarcity episodes.

Second, France is a structurally special market. It operates roughly 63 gigawatts of nuclear capacity, which represents around 70% of annual electricity output. And because French residential heating is predominantly electric, the system has a thermosensitivity of approximately 2.4 gigawatts per degree Celsius in winter — the highest in Europe. That means a cold snap of five degrees adds around 12 gigawatts of demand almost instantly.

Third, the economic stakes of accurate forecasting are very real. A one euro per megawatt-hour improvement in forecast accuracy, on a 100-megawatt portfolio, is worth roughly 876,000 euros per year.

So our starting question is: given all of this, does adding weather data to a forecasting model actually help — or is that signal already captured elsewhere?

---

## SLIDE 3 — RESEARCH QUESTION `[LEO · ~1min30s]`

We decomposed our central question into four sub-questions, shown on this slide.

The first is whether machine learning models — specifically Random Forest and XGBoost — can substantially outperform a naïve statistical benchmark.

The second is whether adding meteorological features produces a statistically significant improvement over a model built on market fundamentals alone.

The third — and this is our core contribution — is whether that value differs between a stable market regime and an extreme-stress regime like the 2022 energy crisis, and why.

And the fourth is whether the predictive improvements translate into real economic value through a day-ahead trading strategy with realistic transaction costs.

These four questions structure the entire thesis, and we will answer each of them in the next twelve slides.

---

## SLIDE 4 — DATA `[LEO · ~1min30s]`

Our dataset spans January 2018 to April 2025, giving us approximately 64,000 hourly observations drawn from four public sources.

From ENTSO-E, we collected hourly day-ahead prices, the load forecast published by RTE, generation by source, and net cross-border flows with Germany, Spain, the UK, and Belgium.

For weather data, we used ERA5 reanalysis from ECMWF — this gives us temperature at 2 metres, wind speed, solar irradiance, and precipitation, all aggregated to a spatial mean over France.

For fuel prices, we included TTF natural gas and ARA coal, which determine the marginal cost of thermal generation.

Finally, we engineered 35 features in total: calendar variables with cyclic encoding, price lags and rolling means, the fundamental variables, cross-border flows, and the weather features including Heating Degree Days and a Weather Stress Index.

One important detail: our data spans three structurally distinct regimes — the pre-crisis period from 2018 to 2021, the 2022 energy crisis where prices briefly exceeded 1,000 euros per megawatt-hour, and the post-crisis normalisation from 2023 onwards. This heterogeneity is actually central to our results, as you will see.

---

## SLIDE 5 — METHODOLOGY `[LEO · ~1min30s]`

Our methodological contribution is a nested four-model ablation design, shown on this slide.

Model A is our naïve benchmark — it predicts the price using the value from exactly 168 hours ago, that is, the same hour of the previous week. This exploits the strong weekly seasonality of electricity prices and is the standard benchmark in the electricity price forecasting literature.

Model B is a Random Forest trained on 27 features — everything except meteorological data. This is our ablation baseline.

Model C is our main model — a Random Forest with all 35 features, including the full weather set.

And Model D is an XGBoost model with the same 35 features, which we include as a more expressive alternative to Random Forest.

The key comparison is C versus B: by holding the architecture constant and changing only the weather features, we isolate the marginal contribution of meteorology to forecast accuracy. This is then evaluated out-of-sample using the Diebold-Mariano test with the Harvey-Leybourne-Newbold correction for finite samples.

---

## SLIDE 6 — RESULTS, STABLE REGIME `[LEO · ~2min]`

Let me now show you the results on our stable test period, which covers May 2024 to April 2025 — 8,640 hourly observations, strictly out-of-sample.

The naïve benchmark achieves a Mean Absolute Error of 33.09 euros per megawatt-hour and an R-squared of 0.16. Already this is a strong signal that the market is somewhat predictable — but 33 euros of average error is still economically significant.

Both Random Forest models reduce that error by 49%, bringing the MAE down to around 16.9 euros per megawatt-hour and the R-squared up to 0.776. This improvement is statistically massive — as we will see on the next slide.

Now look at the critical comparison: Model C, with weather, achieves a MAE of 16.94. Model B, without weather, achieves 16.95. The difference is one hundredth of a euro per megawatt-hour. Essentially zero.

XGBoost, with the same 35 features, actually underperforms both Random Forest models, reaching a MAE of 19.14. We will come back to why.

So the first takeaway is clear: machine learning dramatically beats the naïve benchmark. But weather, in this regime, adds almost nothing.

---

## SLIDE 7 — CENTRAL FINDING `[LEO · ~1min30s]`

This slide formalises that observation with a statistical test.

The Diebold-Mariano test comparing Model C to Model B gives a test statistic of minus 0.565 and a p-value of 0.572. We fail to reject the null hypothesis of equal predictive accuracy. Weather is statistically non-significant in the stable regime.

But look at the other comparisons in the table. When we test whether either Random Forest model beats the naïve benchmark, the DM statistic is around minus 53, with a p-value below 0.001. The predictive gain from using machine learning is massive and unambiguous.

And when we test XGBoost against the RF-weather model, the statistic is positive — meaning XGBoost is significantly *worse* than our main model.

So the picture is this: machine learning adds robust value, but weather, in this regime, does not add anything beyond what is already in the load forecast and gas prices.

Why? And is this always true? I will hand over to Lyam, who will explain the mechanism and then show what happens when we stress-test the model in the 2022 crisis.

---

## SLIDE 8 — FEATURE IMPORTANCE `[LYAM · ~1min30s]`

Thank you, Leo. Let me start by showing you where the predictive power actually comes from — and this will explain why weather is redundant in the stable regime.

The feature importance chart from our Random Forest shows that price lags dominate. The 24-hour lag alone contributes 20.7% of total importance. Lags and rolling means together account for approximately 70%.

After that, fuel prices: TTF gas at 8.3%, ARA coal at 6%. Nuclear availability at roughly 2%.

And all weather features combined — temperature, wind speed, solar radiation, Heating Degree Days, the Weather Stress Index — account for approximately 2.2% of total importance.

This is fully consistent with our DM test result. Weather is not a primary driver of French day-ahead prices in the stable regime. Its signal is already captured elsewhere — and the next slide explains exactly how.

---

## SLIDE 9 — INFORMATION REDUNDANCY HYPOTHESIS `[LYAM · ~2min]`

We call this the *Information Redundancy Hypothesis*, and it rests on three mechanisms.

The first is that the load forecast is equivalent to temperature. RTE publishes an hourly load forecast each morning, and that forecast is itself built using temperature data. So when our model receives the load forecast as a feature, it is implicitly conditioning on weather-driven demand. The raw temperature variable adds no new information.

The second is that TTF gas prices are equivalent to European weather. Gas is the marginal fuel in the European power system. When it is cold across the continent — in Germany, France, Spain — demand for gas rises, and TTF prices rise with it. So the TTF series aggregates pan-European weather into a single commodity price. Again, the raw temperature variable is redundant.

The third is nuclear dominance. In France, nuclear capacity explains a large share of price variance independently of weather. The nuclear availability ratio, which captures planned and unplanned outages, is a much stronger direct driver of prices than the temperature on any given day.

Together, these three channels explain why, in normal market conditions, adding weather variables to a model already equipped with the load forecast, TTF, and nuclear availability does not improve its accuracy. The weather signal is already priced in.

But this reasoning relies on these three channels being intact. In 2022, they were not.

---

## SLIDE 10 — ROBUSTNESS — 2022 CRISIS `[LYAM · ~2min]`

To test the regime-dependence of our finding, we retrained the models on data from 2018 to 2021 and tested them on the full year 2022 — a strict temporal split with no data leakage.

The contrast between the two regimes is stark.

In the stable period: DM statistic of minus 0.565, p-value 0.572 — not significant.

In the 2022 crisis: DM statistic of minus 13.27, p-value below 0.001 — three stars.

The sign of the statistic is the same — Model C with weather is better than Model B without — but in the stable regime the effect is too small to distinguish from noise. In the 2022 crisis, it is one of the strongest DM statistics we have seen in this literature.

In terms of raw accuracy, all models degrade sharply in 2022. The naïve MAE goes from 33 to 73 euros. The RF models go from 16.9 to around 66–67 euros. The market was simply very hard to forecast during a geopolitical shock of that magnitude.

But Model C now beats Model B by 0.73 euros per megawatt-hour on average — and that difference, which sounds small, is statistically enormous because it is consistent across nearly 8,760 hours. XGBoost degrades most in the crisis, reaching a MAE of 76 euros — 14% worse than RF — confirming that Random Forest is the more robust architecture under stress.

---

## SLIDE 11 — WHY THE CRISIS BREAKS THE REDUNDANCY `[LYAM · ~2min]`

Why does weather become significant in 2022? The answer maps directly onto the three redundancy channels we just described — all three break down simultaneously.

First, TTF decouples from local weather. The Russian gas supply shock in 2022 drove TTF prices by geopolitics, not by French or European temperature. At its peak, TTF reached over 300 euros per megawatt-hour — not because it was cold, but because there was no Russian gas. The link between temperature and gas prices snapped. So TTF no longer encoded the weather signal, and the raw temperature variable became independently informative.

Second, nuclear constraints amplify the demand-price relationship. In 2022, EDF discovered stress corrosion cracking in cooling circuit welds across much of the fleet. Nuclear availability fell to approximately 40%. With 63 gigawatts of installed capacity, only around 25 gigawatts were available. There was no headroom to absorb cold-driven demand surges. Every degree of cold translated directly into price — with no nuclear buffer.

Third, signal-to-noise rises. When prices swing by hundreds of euros per megawatt-hour, the direct temperature-to-demand effect becomes large enough to be statistically measurable, even on an hourly basis.

The takeaway is that the value of weather in electricity price forecasting is not a fixed property of the model — it is *regime-dependent*. In normal conditions, include the load forecast and TTF and you have already captured the weather. In a crisis where those proxies break, weather becomes a first-order input.

---

## SLIDE 12 — ECONOMIC VALUE — TRADING BACKTEST `[LYAM · ~1min30s]`

The final empirical question is whether these forecast improvements translate into real economic value. We implemented a day-ahead directional trading strategy with a dead-band filter and three transaction cost scenarios. The results I am showing here use the central cost of 0.30 euros per megawatt-hour, which is consistent with the 2024 EPEX fee schedule for a one-megawatt position.

The naïve strategy generates around 85,000 euros per megawatt per year — but with a maximum drawdown of 2,367 euros, and a Calmar ratio of only 37.

Both Random Forest models generate approximately 137,000 euros per megawatt, with a Sharpe ratio of 19.5. More importantly, the maximum drawdown is only 422 euros — more than five times smaller than the naïve strategy. The Calmar ratio for Model C is 328.

XGBoost generates slightly more in raw profit — 142,000 euros — but its maximum drawdown is 1,351 euros, giving a Calmar of only 107. It earns more on average but with much larger risk.

The Random Forest models produce positive Sharpe ratios in every single one of the twelve test months. This is the evidence that the forecasts contain real skill, not just trend-following.

---

## SLIDE 13 — ROBUSTNESS CHECKS `[LYAM · ~45s]`

Before concluding, let me briefly address robustness.

The trading results hold across all three transaction cost scenarios. RF's Sharpe ratio moves from 19.51 at the lowest cost to 19.16 at the highest — a degradation of less than two percent. The strategy generates a positive Sharpe ratio in every single one of the twelve test months, with no month below zero. And the DM non-significance result for the stable regime is confirmed across multiple sub-periods, not just the full test window.

The conclusion is not sensitive to cost assumptions or to the specific time window chosen.

---

## SLIDE 14 — LIMITATIONS & FUTURE WORK `[LYAM · ~1min]`

We are transparent about six limitations of this work.

First, ERA5 is a reanalysis — it gives us perfect hindsight weather. In production, we would only have NWP forecasts with ten to thirty percent error. Our result is therefore an upper bound on weather value; the true production gain would be smaller.

Second, we trained a single batch model. Monthly walk-forward retraining on an expanding window would better reflect live trading conditions.

Third, EUA carbon prices are excluded. While their signal is partially absorbed by TTF and coal, they should be tested explicitly.

Fourth, XGBoost was run with default hyperparameters. Bayesian optimisation could close or reverse the RF versus XGBoost gap.

Fifth, we study France only. Germany, Spain, and the UK have structurally different generation mixes, and the redundancy mechanism may not apply.

Sixth, transaction costs are modelled as a flat fee. Large positions above ten megawatts would face liquidity constraints not captured here.

These limitations do not invalidate our central finding — but they define the perimeter of what we can claim.

---

## SLIDE 15 — CONCLUSIONS `[LYAM · ~2min]`

Let me close with our four main takeaways.

First, machine learning decisively beats the naïve benchmark. Both Random Forest models reduce MAE by 49% and achieve R-squared of 0.78, with p-values below 0.001. This result is robust across all twelve months of the stable test period.

Second — and this is our core contribution — the value of weather is *regime-dependent*. In stable market conditions, weather is statistically non-significant because its signal is already encoded in the load forecast and TTF gas prices. In the 2022 energy crisis, weather becomes highly significant — the same test yields p below 0.001. The information redundancy hypothesis explains when and why this transition occurs.

Third, Random Forest outperforms XGBoost on risk-adjusted metrics. With our configuration, RF achieves a Calmar ratio of 328 versus 107 for XGBoost, and degrades significantly less under the 2022 stress test. We note that with Bayesian hyperparameter optimisation, XGBoost might close this gap — but that is precisely why we recommend ablation-focused comparisons over off-the-shelf algorithm rankings.

Fourth, the forecasts create real economic value. Approximately 137,000 euros per megawatt per year, a drawdown five times smaller than the naïve benchmark, and a positive Sharpe ratio every single month.

In terms of extensions, the most natural next step is a regime-detection layer — using nuclear availability and TTF volatility as real-time signals to decide when to activate the weather features. That would operationalise the regime-dependent finding directly.

Thank you very much for your attention. We are happy to take your questions.

---

## SLIDE 16 — QUESTIONS `[LEO + LYAM · ~30s]`

> *[Sourire, pause de 2-3 secondes. Laisser l'examinateur parler en premier.]*  
> *[Si silence : "Would you like us to start with any particular aspect of the methodology ?"]*

**[Après chaque question, avant de répondre :]**  
*"That's a great question."* → pause 2 secondes → répondre.

---

## NOTES DE TIMING

| Slide | Speaker | Cible |
|-------|---------|-------|
| 1 — Titre | LEO | 30s |
| 2 — Context | LEO | 2min |
| 3 — Research Q | LEO | 1min30 |
| 4 — Data | LEO | 1min30 |
| 5 — Methodology | LEO | 1min30 |
| 6 — Results stable | LEO | 2min |
| 7 — Central finding | LEO | 1min30 |
| 8 — Feature importance | LYAM | 1min30 |
| 9 — Redundancy | LYAM | 2min |
| 10 — Crise 2022 | LYAM | 2min |
| 11 — Pourquoi 2022 | LYAM | 2min |
| 12 — Backtest | LYAM | 1min30 |
| 13 — Robustness | LYAM | 45s |
| 14 — Limitations | LYAM | 1min |
| 15 — Conclusions | LYAM | 2min |
| 16 — Questions | BOTH | 30s |
| **TOTAL** | | **~21min** |

---

## TRANSITIONS CLÉS À RETENIR

**Slide 7 → 8 (LEO passe à LYAM) :**
> *"I will hand over to Lyam, who will explain the mechanism and then show what happens when we stress-test the model in the 2022 crisis."*

**Slide 9 → 10 (LYAM enchaîne) :**
> *"But this reasoning relies on these three channels being intact. In 2022, they were not."*

**Fin slide 13 :**
> *"Thank you very much for your attention. We are happy to take your questions."*

---

*Bon courage demain — vous maîtrisez le sujet mieux que vous ne le pensez.*

# ORAL DEFENSE SCRIPT — WORD FOR WORD
## The Role of Weather in French Day-Ahead Electricity Price Forecasting
**Leo Cambreleng & Lyam Oumedjeber · EDHEC MSc DAAI · June 2026**

> **FORMAT**: 20 min presentation + Q&A · **LEO** slides 1–7 · **LYAM** slides 8–15 · Slide 16: both · Target: **~17 min**

---

## SLIDE 1 — TITLE `[LEO · ~20 s]`

Good morning, Professor. I'm Leo Cambreleng, and with my co-author Lyam Oumedjeber, we're presenting our Master's thesis: "The Role of Weather in French Day-Ahead Electricity Price Forecasting." I'll cover the first half, then hand over to Lyam.

---

## SLIDE 2 — CONTEXT & MOTIVATION `[LEO · ~1 min 30 s]`

Why forecast French electricity prices — and why might weather help?

First: electricity can't be stored. Supply and demand balance every hour, creating extreme volatility — prices can spike to several thousand euros per megawatt-hour.

Second: France is unique. Around 70% of output is nuclear, and because heating is mostly electric, a one-degree temperature drop adds about 2.4 gigawatts of demand — the highest thermosensitivity in Europe.

Third: the stakes are real. A one-euro improvement on a 100-megawatt book is worth about 876,000 euros per year.

The question is: does weather actually help — or is that signal already captured by other variables?

---

## SLIDE 3 — RESEARCH QUESTION `[LEO · ~1 min]`

We structured the thesis around four sub-questions.

Can machine learning beat a naïve statistical benchmark? Do weather features add significant value beyond fundamentals alone? Does that value change between a stable market and a crisis — and why? And do better forecasts translate into real trading profit?

These four questions map directly onto the slides ahead.

---

## SLIDE 4 — DATA `[LEO · ~1 min]`

Our dataset covers January 2018 to April 2025 — around 64,000 hourly observations from four public sources.

ENTSO-E for prices, load forecast, generation, and cross-border flows. ERA5 from ECMWF for weather — temperature, wind, solar, precipitation. TTF gas and ARA coal for fuel prices. In total, 35 engineered features.

One key point: the data spans three very different regimes — pre-crisis, the 2022 energy crisis where prices briefly exceeded 1,000 euros, and post-crisis normalisation. That heterogeneity is central to our results.

---

## SLIDE 5 — METHODOLOGY `[LEO · ~1 min]`

We used a four-model ablation design.

Model A is the naïve benchmark: the price from the same hour last week. Model B is a Random Forest on 27 features — no weather. Model C is our main model: Random Forest with all 35 features including weather. Model D is XGBoost with the same 35 features.

The key comparison is C versus B: same architecture, same data — only the weather features differ. This isolates weather's marginal contribution, tested out-of-sample with the Diebold-Mariano test.

---

## SLIDE 6 — RESULTS — STABLE REGIME `[LEO · ~1 min 30 s]`

On the stable test period — May 2024 to April 2025, 8,640 observations — the naïve benchmark gives a MAE of 33 euros and an R-squared of 0.16.

Both Random Forest models cut that error by 49%, bringing MAE to 16.9 euros and R-squared to 0.78.

Now the critical number: Model C with weather, 16.94. Model B without weather, 16.95. One hundredth of a euro. XGBoost underperforms at 19.14.

Machine learning clearly works — but weather, in this regime, adds essentially nothing.

---

## SLIDE 7 — CENTRAL FINDING `[LEO · ~1 min]`

The statistical test confirms it. The Diebold-Mariano test for C versus B gives p = 0.572 — not significant.

But test either RF model against the naïve benchmark, and the DM statistic is around minus 53, with p below 0.001. Machine learning adds massive, unambiguous value. Weather does not — in this regime.

Why? And is it always the case? I hand over to Lyam.

> *"I'll hand over to Lyam, who will explain the mechanism and show what happens in the 2022 crisis."*

---

## SLIDE 8 — FEATURE IMPORTANCE `[LYAM · ~1 min]`

Thank you, Leo. The feature importance chart explains why weather is redundant.

Price lags dominate: the 24-hour lag alone contributes 21% of importance, and lags plus rolling means account for about 70%. Then fuel prices: TTF gas 8%, ARA coal 6%, nuclear availability 2%. All weather variables combined: just 2.2%.

Weather's signal is already captured by other features — the next slide explains exactly how.

---

## SLIDE 9 — INFORMATION REDUNDANCY HYPOTHESIS `[LYAM · ~1 min 30 s]`

We call this the Information Redundancy Hypothesis. Three channels explain it.

One: the RTE load forecast is built on temperature data. Including it means the model is already conditioning on weather-driven demand. Raw temperature adds nothing new.

Two: TTF gas prices aggregate pan-European heating demand. When it's cold across Europe, gas demand rises, TTF rises. The commodity market does the weather encoding for us.

Three: in France, nuclear availability explains so much price variance that little remains for weather to explain.

These three channels make weather redundant — as long as they're intact. In 2022, they were not.

> *"But this only holds when those three channels are working. In 2022, they all broke at once."*

---

## SLIDE 10 — 2022 CRISIS REVERSAL `[LYAM · ~1 min 30 s]`

We retrained on 2018–2021 and tested on the full year 2022 — a strict temporal split.

The contrast is stark. Stable period: DM statistic minus 0.565, p = 0.572 — not significant. Crisis 2022: DM statistic minus 13.27, p below 0.001.

All models degrade sharply: naïve MAE goes from 33 to 73 euros, RF from 17 to 67. But Model C now beats Model B by 0.73 euros per megawatt-hour — statistically huge because it's consistent across nearly 8,760 hours.

XGBoost degrades the most at 76 euros MAE — confirming RF as the more robust architecture under stress.

---

## SLIDE 11 — WHY THE CRISIS BREAKS REDUNDANCY `[LYAM · ~1 min 30 s]`

All three redundancy channels break simultaneously in 2022.

One: TTF decouples from local weather. The Russian gas shock drove prices by geopolitics, not temperature. TTF no longer encodes the weather signal.

Two: nuclear constraints amplify the temperature effect. With availability at just 40% — around 25 of 63 gigawatts — there was no buffer. Every cold degree fed directly into price.

Three: signal-to-noise rises. When prices swing by hundreds of euros, the direct temperature effect becomes large enough to measure statistically.

Takeaway: weather's value is not fixed — it's regime-dependent.

---

## SLIDE 12 — ECONOMIC VALUE — TRADING BACKTEST `[LYAM · ~1 min]`

Do better forecasts make money? We ran a day-ahead directional strategy at 0.30 euros per megawatt-hour — the central cost scenario.

The naïve strategy earns 85,000 euros per megawatt but with a drawdown of 2,367 euros and a Calmar ratio of 37. Both RF models earn around 137,000, with a drawdown five times smaller — 422 euros — and a Calmar of 328.

XGBoost earns slightly more in gross profit but its drawdown is 1,351 euros, three times RF's. RF posts positive Sharpe ratios every single month. Real skill, not luck.

---

## SLIDE 13 — ROBUSTNESS CHECKS `[LYAM · ~30 s]`

Quick robustness check. RF's Sharpe ratio across the three cost scenarios moves from 19.51 to 19.16 — under two percent degradation. Positive Sharpe every month. The DM non-significance result holds across multiple sub-periods. The conclusions don't depend on our specific cost assumptions.

---

## SLIDE 14 — LIMITATIONS & FUTURE WORK `[LYAM · ~45 s]`

Six honest limitations.

ERA5 is perfect hindsight weather — in production you'd use forecasts with 10–30% error, so our result is an upper bound. We didn't do walk-forward retraining. Carbon prices were excluded. XGBoost wasn't tuned — Bayesian optimisation could close the gap. We only cover France; other markets may differ. And transaction costs are flat — large books would face liquidity constraints.

These define the perimeter of what we can claim.

---

## SLIDE 15 — CONCLUSIONS `[LYAM · ~1 min 30 s]`

Four takeaways.

One: machine learning decisively beats the naïve benchmark — minus 49% MAE, R-squared 0.78, confirmed across all twelve months.

Two: the value of weather is regime-dependent. Non-significant in stable conditions, highly significant in the 2022 crisis. That's our core contribution.

Three: Random Forest outperforms XGBoost on risk-adjusted metrics — Calmar 328 versus 107 — and holds up better under stress.

Four: the forecasts create real economic value — 137,000 euros per megawatt per year, drawdown five times smaller than the benchmark.

The practical implication: don't always include or always exclude weather. Build a regime detector. Thank you very much.

---

## SLIDE 16 — QUESTIONS `[LEO + LYAM · ~20 s]`

> *[Smile. Pause 2–3 seconds. Let the examiner speak first.]*
> *[If silence: "Would you like us to start with any particular aspect of the methodology?"]*

---

## TIMING REFERENCE

| Slide | Speaker | Target |
|-------|---------|--------|
| 1 — Title | LEO | 20 s |
| 2 — Context & Motivation | LEO | 1 min 30 s |
| 3 — Research Question | LEO | 1 min |
| 4 — Data | LEO | 1 min |
| 5 — Methodology | LEO | 1 min |
| 6 — Results stable | LEO | 1 min 30 s |
| 7 — Central Finding | LEO | 1 min |
| 8 — Feature Importance | LYAM | 1 min |
| 9 — Redundancy | LYAM | 1 min 30 s |
| 10 — 2022 Crisis | LYAM | 1 min 30 s |
| 11 — Why 2022 | LYAM | 1 min 30 s |
| 12 — Trading Backtest | LYAM | 1 min |
| 13 — Robustness | LYAM | 30 s |
| 14 — Limitations | LYAM | 45 s |
| 15 — Conclusions | LYAM | 1 min 30 s |
| 16 — Questions | LEO + LYAM | 20 s |
| **TOTAL** | | **~17 min** |

---

## KEY TRANSITIONS

**Slide 7 → 8 (LEO hands to LYAM):**
> *"I'll hand over to Lyam, who will explain the mechanism and show what happens in the 2022 crisis."*

**Slide 9 → 10 (LYAM continues):**
> *"But this only holds when those three channels are working. In 2022, they all broke at once."*

**End of slide 15:**
> *"Thank you very much. We're happy to take your questions."*

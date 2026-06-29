# ORAL DEFENSE SCRIPT
## The Role of Weather in French Day-Ahead Electricity Price Forecasting
**Leo Cambreleng & Lyam Oumedjeber · EDHEC MSc DAAI · June 2026**

> **LEO** slides 1–7 (~7 min)  ·  **LYAM** slides 8–15 (~9 min 40 s)  ·  **BOTH** slide 16  ·  Total: **~17 min**
>
> **Rule:** don't read the slide — comment on it. Guide attention. Add what isn't written.

---

## SLIDE 1 — TITLE `[LEO · 20 s]`

Good morning, Professor.

I'm Leo — with Lyam, we spent the last year asking a very specific question
about the French electricity market.

I'll set it up, and Lyam will deliver the punchline.

---

## SLIDE 2 — CONTEXT & MOTIVATION `[LEO · 1 min 30 s]`

Before we get into the method — why does this market in particular make the question interesting?

Three things make France unusual.

Electricity can't be stored — so every mismatch between supply and demand
shows up immediately in the price.
That's why you get these violent spikes. It's structurally different from any financial market.

France then adds a layer on top of that:
70% of its electricity comes from nuclear.
And most of its heating is electric.
So a cold snap doesn't just increase demand — it hits a system with very little flexibility.
2.4 gigawatts per degree. The highest in Europe.

And the practical reason we care:
a one-euro improvement in forecast accuracy, on a typical trading book,
is close to a million euros a year.

So the question isn't academic — it's: where does that improvement come from?
Is weather one of the answers?

---

## SLIDE 3 — RESEARCH QUESTION `[LEO · 1 min]`

The question in the box is the one we set out to answer.

But we didn't treat it as a yes/no.
We broke it into four parts — you can see them on the slide.

The first two are about forecast accuracy.
The third — does the answer *change* depending on what the market is doing — is the one that turned out to be the most interesting.
The fourth connects it to something concrete: money.

Those four questions are the backbone of everything that follows.

---

## SLIDE 4 — DATA `[LEO · 1 min]`

On the left, the four data sources — nothing exotic, all public.

The interesting one is ERA5.
It's not a weather forecast — it's a reanalysis.
ECMWF takes actual observations and produces the best possible reconstruction of past weather.
We'll come back to why that distinction matters.

The chart on the right is worth a look before we go further.

You can see the three periods clearly:
the flat, predictable pre-crisis years,
the spike in 2022 — prices briefly above 1,000 euros —
and then the return to something closer to normal.

That's not just historical context. It's the core of our experiment.

---

## SLIDE 5 — METHODOLOGY `[LEO · 1 min]`

The design is simple — and deliberately so.

Four models. One controlled comparison.

B and C are identical in every way except one:
C gets the weather variables, B doesn't.

That's the ablation. If C is better than B, weather helped.
If not, the signal was already there without it.

We test that difference formally with the Diebold-Mariano test —
which is the standard in this literature for comparing forecast accuracy.

---

## SLIDE 6 — RESULTS — STABLE REGIME `[LEO · 1 min 30 s]`

Look at the table — specifically the last two rows on the left.

B and C. 16.95 and 16.94.

One hundredth of a euro difference.
On a baseline of 33 euros, that's noise.

Now look at A versus B — 33 down to 16.9.
That's the machine learning contribution. That's real.

XGBoost underperforms here — we'll explain why that's actually informative.

But the headline finding from this slide:
weather, in a calm market, doesn't move the needle.

---

## SLIDE 7 — CENTRAL FINDING `[LEO · 1 min]`

This slide formalises what we just saw.

The top line — p = 0.572.
We fail to reject equal accuracy between C and B.
Weather is not statistically significant in the stable regime.

Now look at the other rows.
When we test ML against the naïve benchmark,
the DM statistic is minus 53. That's enormous — essentially impossible to get by chance.

So we have two very different stories in the same table:
ML adds unambiguous value. Weather, right now, does not.

The natural question is: is this always true?

> *"Lyam will take it from here."*

---

## SLIDE 8 — FEATURE IMPORTANCE `[LYAM · 1 min]`

Thanks Leo. This chart answers the "why" before I even state it.

Look at the bar lengths.

The top five features are all price lags and rolling averages.
Price history explains about 70% of what the model knows.

Weather sits at the bottom.
All variables combined — temperature, wind, solar, everything — about 2%.

That's not a modelling flaw. It's telling us something real:
weather's signal is already in the data — just encoded differently.

---

## SLIDE 9 — INFORMATION REDUNDANCY HYPOTHESIS `[LYAM · 1 min 30 s]`

Three boxes on this slide — three reasons why weather is redundant.

The first one is the most elegant.
RTE publishes a load forecast every morning.
That forecast is built using temperature data.
So when our model sees the load forecast, it's already implicitly seeing the weather.
Adding raw temperature on top of that adds nothing — it's the same information twice.

The second: TTF gas prices.
Gas is the marginal fuel in Europe.
When it's cold everywhere, gas demand rises, TTF rises.
The commodity market has already aggregated the weather signal for us.

The third: nuclear dominates the French price so much
that weather effects are small in comparison.

Three channels. All of them make weather redundant.

> *"As long as they hold. In 2022, none of them did."*

---

## SLIDE 10 — 2022 CRISIS REVERSAL `[LYAM · 1 min 30 s]`

The two cards at the top of the slide tell the whole story.

Same test. Same models. Different market.

p = 0.572 becomes p < 0.001.
DM = −0.565 becomes DM = −13.27.

In this literature, a DM statistic of minus 3 or 4 is already considered strong.
Minus 13 is almost never seen.

The chart on the right shows why —
2022 is a completely different distribution.
Models degrade badly across the board — that's expected.
But the gap between C and B, which was invisible in stable conditions,
becomes detectable and consistent.

---

## SLIDE 11 — WHY THE CRISIS BREAKS REDUNDANCY `[LYAM · 1 min 30 s]`

Each of the three channels we described — one by one, they break.

TTF first.
In 2022, gas prices were driven by the Russian supply shock, not by temperature.
The link between weather and gas prices snapped.
So TTF stopped encoding the weather signal — and the raw variable had to do it alone.

Nuclear second.
40% availability. That's half the normal level.
In a well-functioning system, nuclear absorbs demand variations.
At 40%, every cold day becomes a price event.

And together, those two effects push price swings into the hundreds of euros —
which is when the temperature signal finally becomes statistically measurable.

The punchline: weather isn't universally useful or useless.
It depends entirely on what the rest of the market is doing.

---

## SLIDE 12 — TRADING BACKTEST `[LYAM · 1 min]`

This slide answers a simple question:
do the forecast improvements translate into anything real?

The table on the left — focus on the drawdown column.

Naïve strategy: 2,367 euros of drawdown per megawatt.
RF models: 422. Five times smaller.

That's the number that matters in practice.
It's not about making more money in absolute terms —
it's about how much you lose when you're wrong.

And the strategy earns positive Sharpe ratios every single month of the test period.
No lucky quarter hiding a bad year. Consistent.

---

## SLIDE 13 — ROBUSTNESS CHECKS `[LYAM · 30 s]`

One concern you might have: do these results depend on the cost assumption we chose?

The table shows the three scenarios.
The Sharpe barely moves — under two percent degradation from cheapest to most expensive.

The conclusion is stable. It's not an artefact of our parameters.

---

## SLIDE 14 — LIMITATIONS `[LYAM · 45 s]`

The most important one is the ERA5 point.

ERA5 gives us the actual weather that happened.
In production, you'd only have a forecast — with real errors.
So our result is a ceiling on how much weather can help.
Even the perfect version of weather barely matters in a stable market.
That actually strengthens the conclusion rather than weakening it.

The other five are genuine limitations —
no walk-forward retraining, no carbon prices, XGBoost not tuned,
one country, flat cost model.
We document them because they define the scope of what we can claim.

---

## SLIDE 15 — CONCLUSIONS `[LYAM · 1 min 30 s]`

Four things to take away.

ML works — clearly and robustly. That's not the surprise.

The surprise is the second point.
Weather matters — but only when the channels that normally encode it break down.
In stable conditions, it's redundant.
In a crisis, it's essential.
That's a regime-dependent finding, and it's the core contribution.

Random Forest holds up better than XGBoost under stress —
which has a direct implication for model selection in volatile markets.

And forecasts create real economic value — not just better numbers on a metric.

The practical takeaway isn't "add weather" or "drop weather."
It's: build a system that knows which regime it's in.

Thank you.

---

## SLIDE 16 — QUESTIONS `[LEO + LYAM · 20 s]`

*[Smile. Wait. Let them speak first.]*

*[If silence: "Would you like to start with any particular aspect?"]*

---

## TIMING REFERENCE

| Slide | Speaker | Target |
|-------|---------|--------|
| 1 — Title | LEO | 20 s |
| 2 — Context | LEO | 1 min 30 s |
| 3 — Research Question | LEO | 1 min |
| 4 — Data | LEO | 1 min |
| 5 — Methodology | LEO | 1 min |
| 6 — Results stable | LEO | 1 min 30 s |
| 7 — Central Finding | LEO | 1 min |
| 8 — Feature Importance | LYAM | 1 min |
| 9 — Redundancy | LYAM | 1 min 30 s |
| 10 — 2022 Crisis | LYAM | 1 min 30 s |
| 11 — Why 2022 | LYAM | 1 min 30 s |
| 12 — Trading | LYAM | 1 min |
| 13 — Robustness | LYAM | 30 s |
| 14 — Limitations | LYAM | 45 s |
| 15 — Conclusions | LYAM | 1 min 30 s |
| 16 — Questions | BOTH | 20 s |
| **TOTAL** | | **~17 min** |

---

## KEY TRANSITIONS

**Slide 7 → 8:** *"Lyam will take it from here."*

**Slide 9 → 10:** *"As long as they hold. In 2022, none of them did."*

**End of slide 15:** *"Thank you."*

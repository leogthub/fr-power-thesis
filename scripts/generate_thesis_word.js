const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageBreak, LevelFormat, TableOfContents,
  convertInchesToTwip
} = require("docx");
const fs = require("fs");
const path = require("path");

const FIGURES = path.join(__dirname, "../outputs/figures");
const OUT     = path.join(__dirname, "../outputs/Thesis_FR_Power_Prices.docx");

// ── Helpers ────────────────────────────────────────────────────────────────

const FONT = "Calibri";
const SZ_BODY   = 24;   // 12pt
const SZ_H1     = 32;   // 16pt
const SZ_H2     = 28;   // 14pt
const SZ_H3     = 26;   // 13pt
const SZ_SMALL  = 20;   // 10pt
const SZ_CAPTION= 20;   // 10pt

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, color: "000000", size: SZ_BODY, ...opts });
}

function p(text, opts = {}) {
  const { align, before = 100, after = 100, indent, children, bold, italic, size } = opts;
  return new Paragraph({
    spacing: { before, after, line: 276, lineRule: "auto" },
    alignment: align || AlignmentType.JUSTIFIED,
    indent,
    children: children || [run(text, { bold: bold || false, italics: italic || false, size: size || SZ_BODY })]
  });
}

function h1(text) {
  return new Paragraph({
    spacing: { before: 400, after: 160 },
    children: [new TextRun({ text, font: FONT, color: "000000", size: SZ_H1, bold: true })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, font: FONT, color: "000000", size: SZ_H2, bold: true })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, font: FONT, color: "000000", size: SZ_H3, bold: true })]
  });
}

function blank(n = 1) {
  return Array.from({ length: n }, () => new Paragraph({
    spacing: { before: 0, after: 0 },
    children: [run("")]
  }));
}

function pb() {
  return new Paragraph({ children: [new TextRun({ text: "", break: 1 }), new PageBreak()] });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 60, after: 60 },
    children: [run(text)]
  });
}

function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    spacing: { before: 80, after: 80 },
    children: [run(text)]
  });
}

function img(filename, widthPx = 500, heightPx = 280) {
  const fp = path.join(FIGURES, filename);
  if (!fs.existsSync(fp)) {
    console.warn(`  WARN: figure not found: ${filename}`);
    return p(`[Figure: ${filename}]`, { italic: true, align: AlignmentType.CENTER });
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 60 },
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(fp),
      transformation: { width: widthPx, height: heightPx },
      altText: { title: filename, description: filename, name: filename }
    })]
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 180 },
    children: [run(text, { italics: true, size: SZ_CAPTION })]
  });
}

function hr() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "000000" } },
    children: [run("")]
  });
}

// Simple table builder — all black, no shading
const brd = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: brd, bottom: brd, left: brd, right: brd };

function makeTable(rows, colWidths) {
  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((cells, ri) => new TableRow({
      children: cells.map((text, ci) => new TableCell({
        borders,
        width: { size: colWidths[ci], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: ci > 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: [run(String(text), { bold: ri === 0, size: SZ_SMALL })]
        })]
      }))
    }))
  });
}

function sp(before = 120, after = 120) {
  return new Paragraph({ spacing: { before, after }, children: [run("")] });
}

// ── Document ───────────────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: SZ_BODY, color: "000000" } }
    }
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
      }
    },
    children: [

      // ═══════════════════════════════════════════════
      // TITLE PAGE
      // ═══════════════════════════════════════════════
      ...blank(4),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        children: [run("The Role of Weather in French Day-Ahead Electricity Price Forecasting: A Random Forest Approach", { bold: true, size: 36 })]
      }),
      hr(),
      ...blank(1),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: 80 },
        children: [run("Master's Thesis — MSc Data Analysis & AI", { size: 26 })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [run("EDHEC Business School", { size: 26 })]
      }),
      ...blank(2),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [run("Leo Camberleng", { bold: true, size: 28 })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [run("Lyam Oumedjeber", { bold: true, size: 28 })]
      }),
      ...blank(2),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [run("May 2026", { size: 24 })]
      }),
      hr(),
      pb(),

      // ═══════════════════════════════════════════════
      // ABSTRACT
      // ═══════════════════════════════════════════════
      h1("Abstract"),
      p("This thesis investigates whether weather variables significantly improve day-ahead electricity price forecasting for the French market (EPEX SPOT France) using machine learning models. We build a comprehensive dataset spanning January 2018 to April 2025, combining hourly price and generation data from the ENTSO-E Transparency Platform with ERA5 reanalysis weather data, cross-border flow data, and fuel price series (TTF natural gas, ARA coal, EU ETS carbon allowances)."),
      p("Four models are compared: a naive lag-168h benchmark (Model A), a Random Forest without weather features (Model B), a Random Forest with full weather features (Model C), and an XGBoost model with weather features (Model D). Models are evaluated on a strict out-of-sample test set covering May 2024 to April 2025 (8,640 hours)."),
      p("Our central finding is that weather features do not significantly improve forecast accuracy over the enriched fundamental model, as confirmed by a Diebold-Mariano test (p = 0.572). Both Random Forest models achieve a Mean Absolute Error of approximately 16.9 EUR/MWh and R2 = 0.776, representing a 49% reduction in error over the naive benchmark. A day-ahead directional trading backtest confirms the economic value of these forecasts, with RF models generating net profits of approximately 136,700 EUR/MW over the test period with a maximum drawdown ten times smaller than the naive strategy."),
      sp(60, 60),
      p("Keywords: electricity price forecasting, day-ahead market, EPEX SPOT France, Random Forest, XGBoost, weather features, Diebold-Mariano test, feature engineering, trading backtest.", { italic: true }),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 1 — INTRODUCTION
      // ═══════════════════════════════════════════════
      h1("1. Introduction"),
      h2("1.1 Context and Motivation"),
      p("Electricity markets have undergone a fundamental transformation over the past three decades. The progressive liberalisation of European energy markets, accelerated by successive EU directives from 1996 onwards, replaced vertically integrated state monopolies with competitive wholesale markets in which prices are determined by the interaction of supply and demand. In France, the day-ahead electricity market operates through EPEX SPOT, where participants submit bids and offers for each hour of the following day, with a single clearing price determined by a batch auction that closes at 12:00 CET."),
      p("The resulting price series is among the most complex in quantitative finance. Unlike commodity prices, electricity cannot be economically stored at scale, which forces supply and demand into instantaneous equilibrium at every moment. The consequence is a price process characterised by extreme volatility, sharp intraday seasonality, strong weekly periodicity, and irregular price spikes that can reach several thousand euros per megawatt-hour during scarcity episodes. The French market adds a layer of specificity: with approximately 63 GW of installed nuclear capacity representing roughly 70% of annual electricity production (RTE, 2023), the availability of the nuclear fleet is a primary determinant of price levels. In parallel, France's residential heating stock is predominantly electric, creating a thermosensitivity of approximately 2.4 GW per degree Celsius in winter (RTE, 2024) — the highest in Europe."),
      p("Accurate forecasting of day-ahead prices is therefore not merely an academic exercise. It is a critical input for energy suppliers setting customer tariffs, for industrial consumers optimising consumption schedules, for renewables producers valuing flexibility options, and for traders positioning in forward markets. A one euro per megawatt-hour improvement in forecast accuracy over a 100 MW portfolio translates to approximately 876,000 EUR in annual value capture."),
      h2("1.2 Research Question"),
      p("This thesis addresses the following central question:"),
      p("Do weather variables significantly improve the accuracy of day-ahead electricity price forecasts for the French market, and if so, what is the economic value of that improvement?", { italic: true, indent: { left: 720, right: 720 } }),
      p("This question is decomposed into three sub-questions:"),
      numbered("Can machine learning models (Random Forest, XGBoost) substantially outperform a naive statistical benchmark on French day-ahead prices?"),
      numbered("Does the addition of meteorological features produce a statistically significant improvement over a model built on market fundamentals alone?"),
      numbered("Can the predictive improvement of these models be translated into economically meaningful returns through a day-ahead directional trading strategy?"),
      h2("1.3 Contributions"),
      p("This thesis makes three contributions to the electricity price forecasting literature:"),
      bullet("Comprehensive French dataset: a reproducible pipeline assembling ENTSO-E generation and load data, ERA5 reanalysis weather data, cross-border flow data, and fuel price series spanning 2018-2025."),
      bullet("Rigorous ablation study: a four-model experiment (A/B/C/D) that isolates the marginal contribution of weather features using a Diebold-Mariano test with Harvey-Leybourne-Newbold correction."),
      bullet("Day-ahead executable trading backtest: a directional strategy grounded in the actual mechanics of the EPEX SPOT day-ahead auction, with three transaction cost scenarios drawn from the 2024 EPEX fee schedule."),
      h2("1.4 Structure"),
      p("Chapter 2 reviews the literature on electricity price forecasting. Chapter 3 describes the data and feature engineering process. Chapter 4 presents modelling results and statistical tests. Chapter 5 introduces the trading backtest. Chapter 6 discusses findings and limitations. Chapter 7 concludes. All code is available at https://github.com/leogthub/fr-power-thesis."),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 2 — LITERATURE
      // ═══════════════════════════════════════════════
      h1("2. Literature Review"),
      h2("2.1 Overview of Electricity Price Forecasting"),
      p("The electricity price forecasting literature has expanded rapidly since market liberalisation in the 1990s. Weron (2014) provides the definitive survey, classifying approaches into five families: multi-agent models, fundamental models, reduced-form stochastic models, statistical models, and computational intelligence models. More recently, Nowotarski and Weron (2018) document the shift towards probabilistic forecasting and machine learning methods. For the purposes of this thesis, we focus on short-term point forecasting at the day-ahead horizon, which remains the most practically relevant setting for market participants."),
      h2("2.2 Statistical and Econometric Approaches"),
      p("Early work on day-ahead price forecasting relied on time-series models adapted from financial econometrics. Autoregressive models — in particular ARIMA and its seasonal extensions — capture the strong autocorrelation structure of electricity prices but struggle to model the non-linear interactions between price and its fundamental drivers. GARCH-family models address volatility clustering but are ill-suited to the extreme price spikes that characterise electricity markets."),
      p("Factor models represent a middle ground. Maciejowska et al. (2016) apply Factor Quantile Regression Averaging to day-ahead prices across European markets. Uniejewski et al. (2019) propose automated variable selection using LASSO regularisation applied to a rich set of price lags and calendar variables, achieving competitive performance relative to more complex non-linear models. This result — that well-specified linear models with price lags can be surprisingly competitive — is relevant to our own findings."),
      h2("2.3 Machine Learning Approaches"),
      p("The application of machine learning to electricity price forecasting has accelerated since the mid-2010s. Lago et al. (2018) offer a comprehensive empirical comparison of deep learning architectures (LSTM, CNN) against traditional benchmarks on European day-ahead markets."),
      p("Tree-based ensemble methods — Random Forests (Breiman, 2001) and Gradient Boosted Trees including XGBoost (Chen and Guestrin, 2016) — have proven particularly well-suited to the electricity price forecasting problem. They handle non-linear interactions between weather, calendar, and fundamental variables without requiring manual feature transformations; are robust to the outlier prices common in electricity markets; and provide interpretable feature importance measures highly valued by market participants."),
      p("Ziel and Weron (2018) compare univariate and multivariate forecasting frameworks on European day-ahead markets and find that multivariate models incorporating load, generation, and cross-border flows outperform univariate price-only models. This supports the use of fundamental variables in our feature matrix."),
      h2("2.4 Weather as a Driver of Electricity Prices"),
      p("The causal pathway from weather to electricity prices is well-established. Temperature affects demand through heating and cooling. Wind speed determines wind generator output. Solar irradiance compresses midday prices during high-penetration periods. Precipitation affects hydro reservoir levels."),
      p("In the French market specifically, thermosensitivity is unusually high by European standards. RTE (2024) reports that each degree Celsius below 17 degrees C increases peak demand by approximately 2.4 GW — a consequence of the high penetration of electric heating (over 30% of French households). This motivates the use of 17 degrees C as the Heating Degree Day threshold throughout this thesis, consistent with RTE's own operational methodology."),
      p("The Weather Stress Index introduced in this thesis captures the interaction between cold temperatures and low wind speed — a combination that simultaneously increases demand and reduces renewable generation, historically associated with the most extreme price spikes in France."),
      h2("2.5 Diebold-Mariano Testing in Energy Forecasting"),
      p("Comparing forecast accuracy requires careful statistical treatment. The Diebold-Mariano test (Diebold and Mariano, 1995) provides a formal framework for testing equal predictive accuracy, accounting for autocorrelation in the loss differential series via a Newey-West long-run variance estimator. For small samples, Harvey et al. (1997) propose a correction factor that adjusts the DM statistic to better approximate a t-distribution. We apply this Harvey-Leybourne-Newbold (HLN) correction throughout our analysis."),
      h2("2.6 Gap in the Literature"),
      p("Despite the growing body of work on machine learning for electricity price forecasting, three gaps remain relevant to this thesis:"),
      numbered("French market specificity: Most studies focus on the German, Spanish, or Nordic markets. The French market's combination of high nuclear share, high thermosensitivity, and interconnection with multiple neighbours creates a distinct price formation mechanism."),
      numbered("Formal ablation of weather features: While weather variables are routinely included, their marginal contribution is rarely isolated using a statistically formal test."),
      numbered("Economic value evaluation: Most studies evaluate models on statistical metrics alone. We extend the analysis to an executable day-ahead trading backtest with realistic transaction costs."),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 3 — DATA
      // ═══════════════════════════════════════════════
      h1("3. Data and Feature Engineering"),
      h2("3.1 Data Sources"),
      h3("3.1.1 ENTSO-E Transparency Platform"),
      p("The European Network of Transmission System Operators for Electricity (ENTSO-E) makes hourly data publicly available through its Transparency Platform. We retrieve the following series for France (bidding zone 10YFR-RTE------C):"),
      bullet("Day-ahead prices (EUR/MWh): the clearing price of the EPEX SPOT France day-ahead auction — our forecasting target."),
      bullet("Actual total load and load forecast (MW): the TSO's hourly load forecast, published the day before delivery."),
      bullet("Actual generation by source (MW): nuclear, gas, run-of-river hydro, reservoir hydro, solar, and onshore wind."),
      bullet("Cross-border physical flows (MW): net hourly flows between France and Germany, Spain, the United Kingdom, and Belgium."),
      h3("3.1.2 ERA5 Reanalysis Weather Data"),
      p("Meteorological variables are sourced from the ERA5 reanalysis dataset produced by ECMWF through the Copernicus Climate Change Service (Hersbach et al., 2020). ERA5 provides hourly estimates on a 0.25 x 0.25 degree grid. We retrieve four variables for metropolitan France and compute spatial means: 2-metre air temperature, 10-metre wind speed, surface solar irradiance, and total precipitation."),
      h3("3.1.3 Fuel Prices"),
      p("Fuel prices include TTF natural gas (EUR/MWh) — the European gas benchmark and primary determinant of marginal cost of gas-fired generation — and ARA coal (EUR/tonne). These are retrieved from public market data sources."),
      h2("3.2 Price Statistics"),
      p("The full price series from 2018 to 2025 exhibits three distinct regimes: a pre-crisis period (2018-2021) with prices typically between 20-80 EUR/MWh; the energy crisis of 2021-2023, during which French prices reached historical extremes above 1,000 EUR/MWh driven by the Russian gas supply disruption and historically low nuclear availability; and a normalisation phase from 2023 onwards."),
      sp(),
      img("price_series_full.png", 520, 240),
      caption("Figure 1 — French day-ahead electricity prices (EUR/MWh), January 2018 - April 2025."),
      sp(),
      img("price_heatmap_hour_month.png", 520, 240),
      caption("Figure 2 — Average hourly price by hour of day and month of year. Winter morning and evening peaks are systematically elevated."),
      h2("3.3 Feature Engineering"),
      p("From the raw data sources, we construct 35 features organised into five categories."),
      h3("3.3.1 Calendar Features"),
      p("The hour of day and month of year are encoded using sine/cosine transformations to preserve their cyclic nature: hour_sin = sin(2*pi*h/24), hour_cos = cos(2*pi*h/24). An is_weekend binary indicator captures the systematic weekday/weekend price differential."),
      h3("3.3.2 Price Lags"),
      p("We include lags at 24h, 48h, and 168h (one week), as well as 24-hour and 168-hour rolling means. The 168-hour lag (same hour, previous week) is particularly important, as weekly seasonality in electricity prices is stronger than daily seasonality."),
      h3("3.3.3 Weather-Derived Features"),
      p("Three derived features are specifically motivated by the French market context:"),
      bullet("Heating Degree Days (HDD): max(0, 17 - T), where T is the 2m temperature. The 17 degrees C threshold is the RTE standard for French demand modelling, reflecting the predominantly electric heating stock."),
      bullet("Wind Power Proxy: v^3, where v is the 10m wind speed. Wind turbine power output is proportional to the cube of wind speed in the operational range."),
      bullet("Weather Stress Index (WSI): HDD / wind_speed. High when temperatures are cold AND wind is weak — the combination historically associated with extreme French price spikes."),
      h3("3.3.4 Nuclear Availability Ratio"),
      p("nuclear_avail = gen_nuclear / 63,000 MW. This ratio captures the fraction of France's installed nuclear fleet currently online. A ratio below 0.7 signals significant outage periods and is historically associated with higher prices."),
      sp(),
      img("price_vs_nuclear.png", 520, 240),
      caption("Figure 3 — Monthly average day-ahead price vs. nuclear availability ratio (2018-2025). Negative correlation is strongest during the 2022 outage crisis."),
      sp(),
      img("price_vs_temperature.png", 520, 240),
      caption("Figure 4 — Day-ahead price vs. 2m temperature. The U-shaped relationship motivates the use of HDD/CDD rather than raw temperature."),
      h2("3.4 Train / Test Split"),
      p("A strict temporal split prevents data leakage. The model is trained on January 2018 to April 2024 and evaluated on the out-of-sample test set from May 2024 to April 2025 (12 months, 8,640 hours). No information from the test period is used in feature construction, model selection, or hyperparameter tuning."),
      sp(),
      makeTable([
        ["Set", "Period", "Hours", "Share"],
        ["Training", "Jan 2018 - Apr 2024", "55,248", "86.4%"],
        ["Test (out-of-sample)", "May 2024 - Apr 2025", "8,640", "13.6%"],
        ["Total", "Jan 2018 - Apr 2025", "63,888", "100%"],
      ], [2800, 2800, 1500, 1000]),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 4 — MODELLING
      // ═══════════════════════════════════════════════
      h1("4. Modelling"),
      h2("4.1 Model Architecture"),
      p("We estimate four models forming a nested hierarchy:"),
      bullet("Model A — Naive benchmark: predicted price = price from same hour previous week (lag-168h). Captures weekly seasonality and serves as the minimum performance bar."),
      bullet("Model B — Random Forest without weather: 27 features including calendar, price lags, generation fundamentals, cross-border flows, and fuel prices. All ERA5 variables are excluded. Serves as the ablation baseline."),
      bullet("Model C — Random Forest with weather (main model): identical RF trained on the full 35-feature matrix."),
      bullet("Model D — XGBoost with weather: gradient-boosted tree model on the same 35 features as Model C."),
      h2("4.2 Model Specifications"),
      h3("4.2.1 Random Forest"),
      p("A Random Forest (Breiman, 2001) is an ensemble of 500 decision trees, each trained on a bootstrap sample with a random subset of sqrt(p) features considered at each split. Hyperparameters: 500 trees, maximum depth 10, minimum 5 observations per leaf. The ensemble prediction is the average across all trees. Tree structure naturally captures threshold non-linearities (e.g., the effect of temperature below the 17 degrees C heating threshold)."),
      h3("4.2.2 XGBoost"),
      p("XGBoost (Chen and Guestrin, 2016) builds 500 trees sequentially, each correcting residuals of the previous ensemble. Hyperparameters: learning rate 0.05, maximum depth 6, subsample 0.8, column subsample 0.8. L1 and L2 regularisation terms prevent overfitting."),
      h2("4.3 Evaluation Metrics"),
      p("Five metrics are reported: MAE (primary, robust to extreme prices), RMSE (penalises large errors), sMAPE (robust to near-zero and negative prices), R2, and Hit Ratio (directional accuracy). MAPE is deliberately excluded as it produces unbounded values with near-zero and negative electricity prices."),
      h2("4.4 Results"),
      sp(),
      makeTable([
        ["Model", "MAE (EUR/MWh)", "RMSE (EUR/MWh)", "sMAPE (%)", "R2", "Hit (%)"],
        ["A - Naive lag-168h",        "33.09", "44.23", "70.9", "0.160", "59.6"],
        ["B - RF without weather",    "16.95", "22.88", "44.2", "0.775", "65.5"],
        ["C - RF with weather",       "16.94", "22.85", "44.2", "0.776", "65.4"],
        ["D - XGBoost with weather",  "19.14", "28.18", "45.4", "0.659", "65.8"],
      ], [2500, 1400, 1500, 1200, 900, 900]),
      sp(60,60),
      p("Table 1 — Out-of-sample forecast performance (May 2024 - April 2025, n = 8,640 hours).", { italic: true, align: AlignmentType.CENTER }),
      sp(),
      p("Key findings: (1) Both Random Forest models reduce MAE by 49% relative to the naive benchmark and increase R2 from 0.160 to 0.776. (2) The addition of weather features (C vs B) produces negligible improvement: MAE decreases by only 0.01 EUR/MWh. (3) XGBoost underperforms both RF models with default hyperparameters (MAE 19.14 vs 16.94 EUR/MWh)."),
      sp(),
      img("model_comparison.png", 520, 200),
      caption("Figure 5 — Out-of-sample performance metrics for all four models."),
      sp(),
      img("forecast_vs_actual.png", 520, 210),
      caption("Figure 6 — Forecast vs actual prices for the first month of the test set (May 2024). Blue = actual, red dashed = RF with weather (C), grey dotted = naive (A)."),
      sp(),
      img("error_distribution.png", 520, 200),
      caption("Figure 7 — Hourly forecast error distribution (left) and actual vs. predicted scatter (right)."),
      sp(),
      img("backtest_mae_by_month.png", 520, 190),
      caption("Figure 8 — MAE by month over the test period. Higher errors in autumn-winter reflect greater price volatility."),
      h2("4.5 Diebold-Mariano Tests"),
      p("We apply the Diebold-Mariano test (1995) with Harvey-Leybourne-Newbold correction (1997) to four model pairs. A negative DM statistic indicates Model 1 is more accurate than Model 2."),
      sp(),
      makeTable([
        ["Comparison", "DM Stat.", "p-value", "Significance"],
        ["C vs B: does weather add value?",       "-0.565", "0.572", "n.s."],
        ["C vs A: does RF-weather beat naive?",   "-53.53", "<0.001", "***"],
        ["B vs A: does RF-no-weather beat naive?","-53.44", "<0.001", "***"],
        ["D vs C: XGBoost vs RF-weather?",        "+10.88", "<0.001", "***"],
      ], [3200, 1200, 1200, 1500]),
      sp(60,60),
      p("Table 2 — Diebold-Mariano tests. *** p<0.01; n.s. not significant.", { italic: true, align: AlignmentType.CENTER }),
      sp(),
      p("Both RF models significantly outperform the naive benchmark (p < 0.001). The addition of weather features is NOT statistically significant (p = 0.572) — the central empirical finding of this thesis. XGBoost significantly underperforms RF-weather (p < 0.001)."),
      h2("4.6 Feature Importance"),
      sp(),
      img("feature_importance_rf.png", 480, 310),
      caption("Figure 9 — Top 20 features by Mean Decrease in Impurity (RF with weather, Model C). Red = weather features, blue = other features."),
      sp(),
      p("Price lags dominate the ranking, with the 168-hour lag, 168-hour rolling mean, and 24-hour rolling mean occupying top positions. Among fundamentals, the TTF gas price and load forecast rank prominently. Nuclear availability ratio also ranks highly. Among weather features, temperature and wind speed appear but with lower importance than price lags or TTF — consistent with the DM test result."),
      h2("4.7 Explanation of the Central Result"),
      p("Three complementary explanations account for the non-significance of weather:"),
      numbered("Load forecast as a weather proxy: RTE's published load forecast already incorporates temperature forecasts. Including this feature implicitly conditions on temperature-driven demand, making the raw temperature variable redundant."),
      numbered("TTF as an energy weather proxy: European gas prices respond to aggregated heating demand across all gas-consuming countries. A model including TTF already captures a pan-European temperature signal through the commodity market."),
      numbered("Nuclear dominance: the nuclear availability ratio explains a large share of price variance independently of weather, leaving little residual for meteorological variables to explain."),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 5 — BACKTEST
      // ═══════════════════════════════════════════════
      h1("5. Day-Ahead Trading Backtest"),
      h2("5.1 Motivation"),
      p("Statistical metrics quantify forecast accuracy but do not directly measure economic value. A model with lower MAE generates higher trading profits only if its accuracy improvements correspond to better directional predictions. This chapter bridges the gap by embedding the four model forecasts in a day-ahead directional trading strategy."),
      h2("5.2 Strategy Design"),
      h3("5.2.1 Market Mechanism"),
      p("The EPEX SPOT France day-ahead market is a uniform-price batch auction. Each day at 12:00 CET, participants submit supply and demand bids for each hour of D+1. The market operator computes the clearing price for each of the 24 hours simultaneously. This means the signal for each hour of D+1 must be computed solely from information available before 12:00 CET on day D — in particular, all 24 hourly prices of day D are available at gate closure."),
      h3("5.2.2 Signal Generation"),
      p("For each hour h of day D+1, the trading signal is:"),
      p("signal(h, D+1) = sign( predicted(h, D+1) - actual(h, D) )", { indent: { left: 720 }, bold: true }),
      p("where predicted(h, D+1) is the model forecast and actual(h, D) is today's realised price for the same hour. This formulation is executable: actual(h, D) is known at gate closure. The signal is +1 (LONG, expect tomorrow's price to be higher) or -1 (SHORT, expect it to be lower), with 1 MW position throughout."),
      p("Dead-band filter: if the predicted day-on-day move is below 2 EUR/MWh, the signal is set to zero (flat position). This avoids executing trades where expected profit does not cover transaction costs."),
      h3("5.2.3 Profit and Loss"),
      p("Gross P&L(h, D+1) = signal(h, D+1) x [actual(h, D+1) - actual(h, D)]"),
      p("Net P&L(h) = Gross P&L(h) - cost x |signal(h)|"),
      p("Costs apply to every active hour (you submit bids each day for each hour you want to trade — not just when you change direction, consistent with EPEX day-ahead mechanics)."),
      h3("5.2.4 Transaction Cost Scenarios"),
      sp(),
      makeTable([
        ["Scenario", "Cost (EUR/MWh)", "Rationale"],
        ["Optimistic",   "0.10", "EPEX matching fee only; large well-connected participants"],
        ["Central",      "0.30", "EPEX fee + bid-ask proxy; mid-size participant"],
        ["Pessimistic",  "0.60", "Central + imbalance settlement risk; conservative upper bound"],
      ], [1500, 1500, 5000]),
      sp(60,60),
      p("Table 3 — Transaction cost scenarios (EPEX SPOT France fee schedule 2024).", { italic: true, align: AlignmentType.CENTER }),
      h2("5.3 Risk Metrics"),
      p("Sharpe ratio: aggregated to daily P&L, annualised with sqrt(365) (electricity trades every calendar day). No risk-free rate deducted. Max Drawdown: largest peak-to-trough in cumulative net P&L. Calmar ratio: annualised P&L / |MDD|. Profit Factor: total positive P&L / total absolute negative P&L. Win Rate: percentage of active hours with positive net P&L."),
      h2("5.4 Results — Central Scenario"),
      sp(),
      makeTable([
        ["Model", "Net P&L (EUR/MW)", "Sharpe", "MDD (EUR/MW)", "Calmar", "PF", "Win%"],
        ["A - Naive",        "85,482",  "10.53", "2,367",  "36.6",  "2.66", "62.0%"],
        ["B - RF no wx",     "136,884", "19.44", "433",    "320.6", "7.07", "74.6%"],
        ["C - RF wx",        "136,711", "19.46", "422",    "328.3", "7.09", "74.6%"],
        ["D - XGBoost",      "142,012", "18.78", "1,351",  "106.6", "7.13", "74.9%"],
        ["Long-only",        "-76",     "-0.01", "4,521",  "-0.02", "1.00", "48.2%"],
      ], [1900, 1600, 1000, 1400, 1000, 800, 900]),
      sp(60,60),
      p("Table 4 — Trading backtest results, central scenario (0.30 EUR/MWh per active hour, May 2024 - April 2025).", { italic: true, align: AlignmentType.CENTER }),
      sp(),
      p("Key findings: (1) RF models generate approximately 136,700 EUR/MW in net P&L over 12 months, versus 85,500 for the naive. (2) RF Maximum Drawdown (422 EUR/MW) is 5.6x smaller than the naive (2,367 EUR/MW) and 3x smaller than XGBoost (1,351 EUR/MW). (3) Models B and C produce nearly identical results, confirming the economic counterpart of the non-significant DM test. (4) The long-only reference earns -76 EUR/MW — returns are not explained by a directional market trend but by genuine forecasting skill."),
      sp(),
      img("equity_curves_net.png", 520, 240),
      caption("Figure 10 — Cumulative net P&L (EUR/MW), central cost scenario. RF models accumulate returns steadily with small drawdowns."),
      sp(),
      img("monthly_pnl.png", 520, 200),
      caption("Figure 11 — Monthly net P&L (EUR/MW) for naive and RF with weather strategies."),
      h2("5.5 Robustness Analysis"),
      h3("5.5.1 Monthly Sharpe Ratio"),
      sp(),
      makeTable([
        ["Month", "Naive", "RF no wx", "RF wx", "XGBoost"],
        ["May 2024",  "9.74",  "19.09", "19.31", "25.48"],
        ["Jun 2024",  "11.73", "23.72", "23.83", "22.55"],
        ["Jul 2024",  "17.88", "28.18", "28.28", "29.75"],
        ["Aug 2024",  "7.92",  "16.60", "16.52", "16.99"],
        ["Sep 2024",  "17.15", "27.08", "27.24", "33.97"],
        ["Oct 2024",  "15.41", "23.01", "22.93", "25.94"],
        ["Nov 2024",  "15.92", "17.91", "17.90", "20.31"],
        ["Dec 2024",  "9.98",  "21.94", "21.89", "26.27"],
        ["Jan 2025",  "5.89",  "17.62", "17.66", "20.32"],
        ["Feb 2025",  "3.91",  "21.83", "21.74", "21.87"],
        ["Mar 2025",  "8.21",  "17.39", "17.59", "22.91"],
        ["Apr 2025",  "12.38", "16.80", "16.80", "-0.99"],
      ], [1600, 1200, 1400, 1400, 1400]),
      sp(60,60),
      p("Table 5 — Monthly Sharpe ratio (daily, annualised sqrt(365)), central cost scenario.", { italic: true, align: AlignmentType.CENTER }),
      sp(),
      p("RF models (B and C) record positive Sharpe in all 12 months. No single month drives the annual result. XGBoost records negative Sharpe in April 2025, suggesting higher sensitivity to market regime changes."),
      sp(),
      img("rolling_sharpe.png", 520, 210),
      caption("Figure 12 — 30-day rolling Sharpe ratio (daily net P&L, annualised sqrt(365)) for each model."),
      h3("5.5.2 Cost Sensitivity"),
      sp(),
      makeTable([
        ["Model", "Optimistic (0.10)", "Central (0.30)", "Pessimistic (0.60)"],
        ["A - Naive",      "10.73", "10.53", "10.23"],
        ["B - RF no wx",   "19.64", "19.44", "19.14"],
        ["C - RF wx",      "19.66", "19.46", "19.16"],
        ["D - XGBoost",    "18.98", "18.78", "18.47"],
      ], [2000, 1700, 1700, 2000]),
      sp(60,60),
      p("Table 6 — Net Sharpe ratio by transaction cost scenario. RF performance is robust across all cost assumptions.", { italic: true, align: AlignmentType.CENTER }),
      sp(),
      img("cost_sensitivity.png", 520, 210),
      caption("Figure 13 — Net P&L and Sharpe by model and transaction cost scenario."),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 6 — DISCUSSION
      // ═══════════════════════════════════════════════
      h1("6. Discussion"),
      h2("6.1 The Central Finding: Weather Is Not Significant"),
      p("The non-significance of weather features (p = 0.572, DM test) is the thesis's primary empirical contribution. This result runs counter to the intuitive expectation that temperature, wind, and solar radiation — which directly drive electricity demand and renewable generation supply — should be informative for price prediction."),
      h3("6.1.1 The Information Redundancy Hypothesis"),
      p("We argue that the apparent paradox resolves once we recognise that weather information is already embedded in other features:"),
      numbered("Load forecast as a temperature proxy: RTE's published load forecast already incorporates temperature forecasts. Including this feature implicitly conditions on temperature-adjusted demand, making the raw temperature variable redundant."),
      numbered("TTF gas price as an energy weather proxy: European gas prices respond to heating demand aggregated across all gas-consuming countries. A model that includes TTF already incorporates a pan-European temperature signal through the commodity market."),
      numbered("Nuclear availability dominance: the nuclear availability ratio explains a disproportionate share of price variance in France. With this feature included, residual variance is already low, leaving little room for weather variables to contribute."),
      h3("6.1.2 Implications for Forecasting Practice"),
      p("Building a rich fundamental model that includes cross-border flows, fuel prices, and nuclear availability may be sufficient to capture the weather signal implicitly in the French market. The additional cost and complexity of integrating real-time weather forecasts may not be justified, provided that these fundamentals are available. This result may not generalise to markets with higher renewable penetration, such as Germany, where the direct impact of weather on supply is less likely to be captured by fuel prices or load forecasts alone."),
      h2("6.2 XGBoost Underperformance"),
      p("XGBoost's sequential boosting mechanism may be more susceptible to the extreme price observations present in the training set (2022 crisis prices above 1,000 EUR/MWh). Each boosting iteration gives elevated weight to poorly-fitted observations; in the presence of genuine outliers driven by extraordinary events, this can lead to overfitting to the crisis regime at the expense of normal-market performance. Systematic hyperparameter optimisation using Bayesian search with a temporal validation strategy could close or reverse the performance gap."),
      h2("6.3 Limitations"),
      numbered("Static model: both models are trained once on 2018-2024 data. In production, periodic retraining (monthly or quarterly) would be needed to adapt to structural market changes."),
      numbered("Price spike forecasting: both models systematically underestimate extreme price events (above 200 EUR/MWh). A two-stage model (spike classifier + regression) or quantile regression could improve performance in the tails."),
      numbered("Negative price handling: the French market has growing frequency of negative prices in spring-summer; the models may not fully capture dynamics below zero."),
      numbered("Single test period: 12 months of one market regime. A rolling out-of-sample evaluation over multiple years would provide more robust assessment."),
      numbered("Perfect execution assumed in backtest: no slippage, no partial fills, 1 MW position size throughout."),
      h2("6.4 Future Research Directions"),
      bullet("Probabilistic forecasting: replace point forecasts with prediction intervals using Quantile Regression Forests or conformal prediction."),
      bullet("Multi-horizon forecasting: extend from day-ahead to 2-day-ahead, requiring recursive or direct strategies."),
      bullet("France vs Germany comparative study: test whether weather features are more significant in Germany where renewable penetration exceeds 50%."),
      bullet("XGBoost hyperparameter optimisation via Bayesian search with temporal cross-validation."),
      bullet("Deep learning: LSTM and Transformer architectures for long-range dependencies."),
      pb(),

      // ═══════════════════════════════════════════════
      // CHAPTER 7 — CONCLUSION
      // ═══════════════════════════════════════════════
      h1("7. Conclusion"),
      h2("7.1 Summary of Findings"),
      p("This thesis investigated whether weather variables significantly improve day-ahead electricity price forecasting for the French market, and whether the resulting forecasts generate economic value through a directional trading strategy. Our analysis leads to four principal conclusions."),
      p("First, machine learning models substantially outperform the naive benchmark. Both Random Forest variants reduce MAE by approximately 49% (from 33.09 to 16.94 EUR/MWh) and increase R2 from 0.16 to 0.78. The improvement is highly statistically significant (p < 0.001) and robust across all 12 months of the test period."),
      p("Second, weather features do not significantly improve forecast accuracy in the French market (DM test p = 0.572). The information redundancy hypothesis — that load forecast, TTF gas price, and nuclear availability already implicitly encode the relevant meteorological information — explains this surprising result. This is the thesis's primary empirical contribution."),
      p("Third, Random Forest outperforms XGBoost with default hyperparameters on a risk-adjusted basis. While XGBoost achieves slightly higher absolute P&L, its Maximum Drawdown is more than three times larger (1,351 vs 422 EUR/MW), yielding a Calmar ratio of 107 versus 320-328 for RF."),
      p("Fourth, the forecasting models generate economically significant value. Under the central transaction cost scenario, RF models generate net P&L of approximately 136,700 EUR/MW over the 12-month test period, with performance robust to cost assumptions (cost drag below 2%) and consistent across all calendar months."),
      h2("7.2 Recommendations"),
      bullet("Use Random Forest as the primary forecasting model for the French day-ahead market, given its superior risk-adjusted performance and interpretable feature importances."),
      bullet("Prioritise fundamental features (load forecast, fuel prices, nuclear availability, cross-border flows) over raw weather variables. Given the ablation result, model development effort is better spent improving fundamental data quality."),
      bullet("Implement periodic retraining (monthly or quarterly) to adapt to structural market changes."),
      bullet("Complement point forecasts with prediction intervals for risk management applications, given the high RMSE relative to MAE during spike periods."),
      h2("7.3 Broader Implications"),
      p("The finding that weather is non-significant in the French day-ahead market reflects a broader insight about information aggregation in commodity markets: when market prices (TTF gas) and operational forecasts (RTE load forecast) already incorporate weather expectations, raw meteorological variables become redundant from a statistical standpoint. This principle — that feature selection should account for information already embedded in price-based variables — is likely applicable beyond the electricity market to other commodity forecasting problems."),
      p("From a market design perspective, this thesis illustrates how the quality of publicly available data (ENTSO-E generation data, RTE load forecasts) can substitute for proprietary weather modelling infrastructure, lowering barriers to participation in day-ahead electricity price forecasting."),
      h2("7.4 Data and Code Availability"),
      p("All code, data pipelines, trained models, and outputs are openly available at: https://github.com/leogthub/fr-power-thesis"),
      pb(),

      // ═══════════════════════════════════════════════
      // REFERENCES
      // ═══════════════════════════════════════════════
      h1("References"),
      p("Breiman, L. (2001). Random Forests. Machine Learning, 45(1), 5-32."),
      p("Chen, T. and Guestrin, C. (2016). XGBoost: A Scalable Tree Boosting System. Proceedings of the 22nd ACM SIGKDD, 785-794."),
      p("Diebold, F.X. and Mariano, R.S. (1995). Comparing predictive accuracy. Journal of Business & Economic Statistics, 13(3), 253-263."),
      p("ENTSO-E (2024). Transparency Platform. https://transparency.entsoe.eu"),
      p("EPEX SPOT (2024). Day-Ahead Market Trading Rules and Fee Schedule. https://www.epexspot.com"),
      p("Fanone, E., Gamba, A. and Prokopczuk, M. (2013). The case of negative day-ahead electricity prices. Energy Economics, 35, 22-34."),
      p("Harvey, D., Leybourne, S. and Newbold, P. (1997). Testing the equality of prediction mean squared errors. International Journal of Forecasting, 13(2), 281-291."),
      p("Hersbach, H. et al. (2020). The ERA5 global reanalysis. Quarterly Journal of the Royal Meteorological Society, 146(730), 1999-2049."),
      p("Lago, J., De Ridder, F. and De Schutter, B. (2018). Forecasting spot electricity prices: Deep learning approaches and empirical comparison of traditional algorithms. Applied Energy, 221, 386-405."),
      p("Maciejowska, K., Nowotarski, J. and Weron, R. (2016). Probabilistic forecasting of electricity spot prices using Factor Quantile Regression Averaging. International Journal of Forecasting, 32(3), 957-965."),
      p("Newey, W.K. and West, K.D. (1987). A simple, positive semi-definite, heteroskedasticity and autocorrelation consistent covariance matrix. Econometrica, 55(3), 703-708."),
      p("Nowotarski, J. and Weron, R. (2018). Recent advances in electricity price forecasting: A review of probabilistic forecasting. Renewable and Sustainable Energy Reviews, 81, 1548-1568."),
      p("RTE (2023). Bilan electrique 2023. Reseau de Transport d'Electricite."),
      p("RTE (2024). Rapport sur la flexibilite et la thermosensibilite de la demande electrique. Reseau de Transport d'Electricite."),
      p("Staffell, I. and Pfenninger, S. (2016). Using bias-corrected reanalysis to simulate current and future wind power output. Energy, 114, 1224-1239."),
      p("Uniejewski, B., Nowotarski, J. and Weron, R. (2019). Automated variable selection and shrinkage for day-ahead electricity price forecasting. Energies, 12(23), 4561."),
      p("Weron, R. (2014). Electricity price forecasting: A review of the state-of-the-art with a look into the future. International Journal of Forecasting, 30(4), 1030-1081."),
      p("Ziel, F. and Weron, R. (2018). Day-ahead electricity price forecasting with high-dimensional structures. Energy Economics, 76, 321-334."),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Thesis Word document saved -> " + OUT);
});

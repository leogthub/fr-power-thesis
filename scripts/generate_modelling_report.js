const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, PageNumber, PageBreak, LevelFormat
} = require("docx");
const fs = require("fs");
const path = require("path");

const FIGURES = path.join(__dirname, "../outputs/figures");
const OUT = path.join(__dirname, "../outputs/Modelling_Report_French_Power_Prices.docx");

function img(filename, width, height) {
  const filepath = path.join(FIGURES, filename);
  if (!fs.existsSync(filepath)) return null;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new ImageRun({
      type: "png",
      data: fs.readFileSync(filepath),
      transformation: { width, height },
      altText: { title: filename, description: filename, name: filename }
    })]
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 200 },
    children: [new TextRun({ text, italics: true, size: 18, color: "555555" })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: "1F3864" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, size: 24, color: "2E5FA3" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 100 },
    children: [new TextRun({ text, size: 22, ...opts })]
  });
}

function rule() {
  return new Paragraph({
    spacing: { before: 100, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" } },
    children: []
  });
}

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function makeTable(rows, colWidths, headerColor = "1F3864") {
  return new Table({
    width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: colWidths,
    rows: rows.map((cells, i) => new TableRow({
      children: cells.map((text, j) => new TableCell({
        borders,
        width: { size: colWidths[j], type: WidthType.DXA },
        shading: {
          fill: i === 0 ? (j === 0 ? headerColor : "2E5FA3") : (i % 2 === 0 ? "F0F4FA" : "FFFFFF"),
          type: ShadingType.CLEAR
        },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        children: [new Paragraph({
          alignment: j > 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: [new TextRun({
            text: String(text),
            size: 20,
            bold: i === 0,
            color: i === 0 ? "FFFFFF" : "333333"
          })]
        })]
      }))
    }))
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E5FA3" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E5FA3" } },
          children: [
            new TextRun({ text: "Modelling Report — French Day-Ahead Power Prices", size: 18, color: "555555" }),
            new TextRun({ text: "\tEDHEC MSc Data Analysis & AI", size: 18, color: "555555" }),
          ],
          tabStops: [{ type: "right", position: 9026 }]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" } },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" }),
            new TextRun({ text: " / ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888" }),
          ]
        })]
      })
    },
    children: [

      // ── TITLE ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 100 },
        children: [new TextRun({ text: "MODELLING REPORT", size: 40, bold: true, color: "1F3864" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Day-Ahead French Power Price Forecasting", size: 28, color: "2E5FA3" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "EDHEC MSc Data Analysis & AI  —  Master's Thesis", size: 20, italics: true, color: "777777" })]
      }),
      rule(),

      // ── 1. METHODOLOGY ──
      h1("1. Methodology"),
      h2("1.1 Problem Formulation"),
      p("The objective is to forecast the French day-ahead electricity price (EPEX SPOT France) for each hour of the next day, using information available before the market gate closure (~12:00 CET). This is a supervised regression problem on a time series with strong seasonal structure."),

      h2("1.2 Feature Engineering"),
      p("28 features were constructed from four data sources:"),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Calendar: hour of day (sin/cos encoded), day of week, month (sin/cos encoded), is_weekend", size: 22 })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Price lags: T-24h, T-48h, T-168h, rolling means (24h, 168h)", size: 22 })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Fundamentals: load forecast, nuclear/wind/solar/gas/hydro generation (MW)", size: 22 })] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 60, after: 200 },
        children: [new TextRun({ text: "Weather (ERA5): temperature, wind speed, solar radiation, precipitation, HDD, CDD, wind power proxy, Weather Stress Index", size: 22 })] }),

      h2("1.3 Train / Test Split"),
      p("A strict temporal split is used to prevent data leakage. The model is trained on 2018–2023 and evaluated on the full year 2024 (out-of-sample)."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      makeTable([
        ["Set", "Period", "Hours"],
        ["Training", "Jan 2018 — Dec 2023", "~52,500"],
        ["Test (out-of-sample)", "Jan 2024 — Dec 2024", "~8,700"],
      ], [3500, 3500, 2026]),

      // ── 2. MODELS ──
      new Paragraph({ children: [new PageBreak()] }),
      h1("2. Models"),
      h2("2.1 Naive Benchmark"),
      p("The naive model predicts the price at hour h on day D+1 as equal to the price at hour h on day D (same-hour previous day). This simple benchmark captures the strong 24h autocorrelation but ignores all other information. It serves as the minimum performance bar any model must beat."),

      h2("2.2 Random Forest"),
      p("A Random Forest regressor (500 trees, max depth 10, min samples leaf 5) is trained on the full feature matrix. Random Forests are well-suited to electricity price forecasting because they handle non-linear interactions between weather and calendar variables without manual specification, are robust to outliers, and provide built-in feature importance measures. The ensemble approach also reduces variance compared to individual decision trees."),

      h2("2.3 XGBoost"),
      p("An XGBoost gradient-boosted tree model (500 estimators, learning rate 0.05, max depth 6, subsample 0.8) is trained as an alternative to the Random Forest. XGBoost typically performs comparably or better on tabular data due to its sequential error-correction mechanism and built-in regularisation (L1/L2)."),

      // ── 3. RESULTS ──
      h1("3. Results"),
      h2("3.1 Test Set Performance"),
      p("All models are evaluated on the 2024 out-of-sample test set. The table below reports Mean Absolute Error (MAE), Root Mean Square Error (RMSE), and R-squared (R²)."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      makeTable([
        ["Model", "MAE (€/MWh)", "RMSE (€/MWh)", "R²"],
        ["Naive (benchmark)", "20.37", "28.37", "0.514"],
        ["Random Forest",     "15.71", "20.80", "0.739"],
        ["XGBoost",           "15.36", "20.11", "0.756"],
      ], [3000, 2000, 2000, 2026]),
      new Paragraph({ spacing: { before: 120, after: 80 }, children: [] }),
      p("Both tree-based models substantially outperform the naive benchmark, reducing MAE by ~25% and improving R² by over 24 percentage points. XGBoost achieves marginally better performance across all metrics (MAE: 15.36 vs 15.71, R²: 0.756 vs 0.739)."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      img("model_comparison.png", 560, 200) || p("[Figure: model_comparison.png]"),
      caption("Figure 1 — MAE, RMSE, and R² comparison across all models on the 2024 test set."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      h2("3.2 Forecast vs Actual"),
      p("Figure 2 shows the first month of the test set. Both models track the general price pattern well, including the intra-week structure and daily peaks. Errors are largest during sudden price spikes driven by unforeseeable events (plant outages, gas supply shocks)."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("forecast_vs_actual.png", 560, 260) || p("[Figure: forecast_vs_actual.png]"),
      caption("Figure 2 — Forecast vs actual prices for the first month of the test set (January 2024)."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.3 Error Distribution"),
      p("Forecast errors are approximately centred at zero, confirming absence of systematic bias. The distribution has heavier tails for extreme price events — consistent with the known difficulty of forecasting price spikes driven by rare supply disruptions. XGBoost and RF exhibit very similar error profiles."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("error_distribution.png", 560, 200) || p("[Figure: error_distribution.png]"),
      caption("Figure 3 — Error distribution (left) and actual vs. predicted scatter plot (right)."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      h2("3.4 Temporal Stability — MAE by Month"),
      p("Figure 4 decomposes test-set MAE by month. Both models show higher errors in winter months (January, February, December), where price volatility is greatest due to heating demand peaks. Performance is more stable in spring and summer. This motivates a potential extension: season-specific models or re-weighting of winter observations."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("backtest_mae_by_month.png", 560, 185) || p("[Figure: backtest_mae_by_month.png]"),
      caption("Figure 4 — MAE by month on the 2024 test set. Higher errors in winter reflect greater price volatility."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.5 Feature Importance"),
      p("Figure 5 shows the top 20 features by Random Forest importance. Price lags dominate (T-24h, T-168h, rolling means), confirming strong auto-regressive structure. Among weather features, temperature (and derived HDD) and wind speed rank highly, validating the thesis hypothesis that weather is a significant driver of French day-ahead prices. The Weather Stress Index (WSI) also appears in the top 20, justifying its engineering."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("feature_importance_rf.png", 480, 310) || p("[Figure: feature_importance_rf.png]"),
      caption("Figure 5 — Top 20 feature importances (Random Forest). Red = weather features, blue = others."),

      // ── 4. DISCUSSION ──
      new Paragraph({ children: [new PageBreak()] }),
      h1("4. Discussion & Limitations"),

      new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: "Price spike forecasting: both models underestimate extreme price events (>200 €/MWh). A dedicated spike classifier or quantile regression extension could improve tail performance.", size: 22 })] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: "MAPE unreliability: MAPE is very high due to near-zero and negative prices in the denominator. MAE and RMSE are the primary metrics for this market.", size: 22 })] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { before: 80, after: 80 },
        children: [new TextRun({ text: "Static model: the model is trained once on 2018-2023. In production, periodic retraining (e.g. monthly) would be needed to adapt to structural market changes.", size: 22 })] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { before: 80, after: 120 },
        children: [new TextRun({ text: "Missing fundamentals: cross-border flows and fuel price futures (TTF gas, EUA carbon) were not included and could further improve accuracy.", size: 22 })] }),

      rule(),
      h1("5. Conclusion"),
      p("Both Random Forest and XGBoost models significantly outperform the naive benchmark on 2024 out-of-sample data, achieving MAE around 15.4–15.7 €/MWh and R² of 0.74–0.76. Weather features — particularly temperature, wind speed, and the Weather Stress Index — contribute meaningfully to forecast accuracy alongside price lags and generation fundamentals. XGBoost is the recommended model for deployment given its marginally superior performance and faster training time."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      rule(),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 120 },
        children: [new TextRun({ text: "Leo Camberleng — EDHEC MSc Data Analysis & AI — May 2026", size: 18, italics: true, color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Report saved -> " + OUT);
});

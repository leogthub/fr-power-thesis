const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak,
  LevelFormat
} = require("docx");
const fs = require("fs");
const path = require("path");

const FIGURES = path.join(__dirname, "../outputs/figures");
const OUT = path.join(__dirname, "../outputs/EDA_Report_French_Power_Prices.docx");

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

function statTable(rows) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [4513, 4513],
    rows: rows.map(([label, value], i) => new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: 4513, type: WidthType.DXA },
          shading: { fill: i === 0 ? "1F3864" : (i % 2 === 0 ? "F0F4FA" : "FFFFFF"), type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          children: [new Paragraph({ children: [new TextRun({
            text: label, size: 20, bold: i === 0, color: i === 0 ? "FFFFFF" : "333333"
          })] })]
        }),
        new TableCell({
          borders,
          width: { size: 4513, type: WidthType.DXA },
          shading: { fill: i === 0 ? "2E5FA3" : (i % 2 === 0 ? "F0F4FA" : "FFFFFF"), type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({
            text: value, size: 20, bold: i === 0, color: i === 0 ? "FFFFFF" : "333333"
          })] })]
        }),
      ]
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
            new TextRun({ text: "EDA Report — French Day-Ahead Power Prices", size: 18, color: "555555" }),
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

      // ── TITLE PAGE BLOCK ──
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 400, after: 100 },
        children: [new TextRun({ text: "EXPLORATORY DATA ANALYSIS", size: 40, bold: true, color: "1F3864" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "French Day-Ahead Power Market  |  2018–2024", size: 28, color: "2E5FA3" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "EDHEC MSc Data Analysis & AI  —  Master’s Thesis", size: 20, italics: true, color: "777777" })]
      }),
      rule(),

      // ── 1. INTRODUCTION ──
      h1("1. Introduction"),
      p("This report summarises the exploratory data analysis (EDA) conducted on the French day-ahead electricity market. The dataset covers hourly observations from January 2018 to December 2024 (61,344 hours), combining four sources:"),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "ENTSO-E Transparency Platform: day-ahead prices, load forecast, generation by fuel type", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "ERA5 (Copernicus / ECMWF): hourly 2m temperature, 10m wind speed, solar radiation, precipitation", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 60, after: 200 },
        children: [new TextRun({ text: "Derived features: Heating Degree Days (HDD), wind power proxy, Weather Stress Index (WSI)", size: 22 })]
      }),

      // ── 2. PRICE STATISTICS ──
      h1("2. Price Dynamics"),
      h2("2.1 Descriptive Statistics"),
      p("The table below summarises the distribution of French day-ahead prices over the full sample period."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      statTable([
        ["Statistic", "Value"],
        ["Mean", "94.50 €/MWh"],
        ["Median", "57.88 €/MWh"],
        ["Std. Deviation", "104.34 €/MWh"],
        ["Minimum", "−134.94 €/MWh"],
        ["Maximum", "2,987.78 €/MWh"],
        ["Negative-price hours", "~2.1% of observations"],
        ["Hours above 200 €/MWh", "~8.4% of observations"],
      ]),
      new Paragraph({ spacing: { before: 120, after: 80 }, children: [] }),
      p("The large gap between mean (94.50) and median (57.88) reflects strong positive skewness driven by crisis episodes in 2021–2022 (gas supply shock, post-COVID demand rebound, nuclear maintenance wave). Negative prices occur mainly during low-demand weekends with high solar and wind output."),

      new Paragraph({ spacing: { before: 160 }, children: [] }),
      img("price_series_full.png", 560, 160) || p("[Figure: price_series_full.png]"),
      caption("Figure 1 — French day-ahead prices 2018–2024 (hourly). The 2021–2022 energy crisis is clearly visible."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      img("price_distribution.png", 560, 200) || p("[Figure: price_distribution.png]"),
      caption("Figure 2 — Price distribution (left) and annual boxplots (right). The 2022 outliers compress the scale."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      // ── 2.2 SEASONALITY ──
      h2("2.2 Seasonality"),
      p("Prices exhibit strong intra-day, intra-week, and seasonal patterns. The heatmap (Figure 4) is particularly informative: morning peaks (7–9h) and evening peaks (18–20h) are consistent year-round, while winter months show structurally higher prices due to heating demand."),

      new Paragraph({ spacing: { before: 100 }, children: [] }),
      img("price_seasonality.png", 560, 175) || p("[Figure: price_seasonality.png]"),
      caption("Figure 3 — Average price by month, day of week, and hour of day."),

      new Paragraph({ spacing: { before: 100 }, children: [] }),
      img("price_heatmap_hour_month.png", 560, 230) || p("[Figure: price_heatmap_hour_month.png]"),
      caption("Figure 4 — Heatmap of average price by hour × month. Dark red = high price, green = low price."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("2.3 Nuclear Generation"),
      p("France is structurally dependent on nuclear power (~70% of production). The 2022 crisis was amplified by an unprecedented nuclear maintenance wave that reduced availability to historical lows, pushing prices to extreme levels."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("price_vs_nuclear.png", 560, 200) || p("[Figure: price_vs_nuclear.png]"),
      caption("Figure 5 — Monthly average price (blue) vs. nuclear output (GW, orange). The inverse relationship is clear in 2022."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      // ── 3. WEATHER ──
      h1("3. Weather Variables & Impact on Prices"),
      h2("3.1 Overview"),
      p("ERA5 reanalysis data provides hourly weather observations spatially averaged over mainland France. Figure 6 shows the monthly evolution of the four weather variables used in the model."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("weather_overview.png", 560, 230) || p("[Figure: weather_overview.png]"),
      caption("Figure 6 — Monthly averages of ERA5 weather variables for France (2018–2024)."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.2 Temperature & Heating Demand"),
      p("Cold temperatures drive up electricity demand through electric heating, a dominant effect in France. The U-shaped relationship in Figure 7 confirms this: both very cold and moderately warm periods see higher prices (cooling demand in summer), with the lowest prices in mild spring/autumn months."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("price_vs_temperature.png", 560, 185) || p("[Figure: price_vs_temperature.png]"),
      caption("Figure 7 — Scatter and binned average price by temperature. U-shaped pattern reflects dual heating/cooling demand."),

      new Paragraph({ spacing: { before: 100 }, children: [] }),
      img("price_vs_hdd.png", 560, 180) || p("[Figure: price_vs_hdd.png]"),
      caption("Figure 8 — Monthly price vs. Heating Degree Days (HDD). High HDD periods (winter) correlate with higher prices."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      h2("3.3 Wind Speed & Renewable Generation"),
      p("Wind generation in France (primarily onshore) can displace gas-fired peakers and lower prices significantly. Low-wind periods combined with cold temperatures create the most extreme price spikes."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("price_vs_wind.png", 560, 185) || p("[Figure: price_vs_wind.png]"),
      caption("Figure 9 — Price vs. wind speed. Higher wind speed is associated with lower prices (merit-order effect)."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.4 Weather Stress Index"),
      p("The Weather Stress Index (WSI) is a composite variable defined as HDD divided by wind speed. It captures simultaneous high-demand and low-renewable-supply conditions — the configuration most likely to produce extreme price spikes. The correlation between WSI and price is 0.42 (Pearson), the highest of all individual weather features."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("weather_stress_index.png", 560, 185) || p("[Figure: weather_stress_index.png]"),
      caption("Figure 10 — Monthly WSI (orange dashed) vs. price (blue). Strong co-movement, particularly in winter 2021–2022."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.5 Correlation Structure"),
      p("The correlation matrix below confirms the expected relationships: nuclear output is negatively correlated with price, wind is negatively correlated, load positively correlated, and the WSI positively correlated."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("correlation_full.png", 480, 340) || p("[Figure: correlation_full.png]"),
      caption("Figure 11 — Correlation matrix across price, weather, and fundamental variables."),

      new Paragraph({ spacing: { before: 160 }, children: [] }),
      rule(),

      // ── 4. KEY FINDINGS ──
      h1("4. Key Findings & Modelling Implications"),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 80, after: 60 },
        children: [new TextRun({ text: "Strong seasonality at hourly, daily, and monthly levels — calendar features (hour, day-of-week, month) are essential inputs.", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Price lags (T−24h, T−48h, T−168h) capture auto-regressive structure and regime persistence.", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "Nuclear availability is the dominant French market fundamental — it should be treated as a primary feature.", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 60 },
        children: [new TextRun({ text: "The Weather Stress Index (WSI = HDD / wind speed) has higher correlation with price than raw temperature or wind alone — justifying its inclusion as an engineered feature.", size: 22 })]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before: 60, after: 120 },
        children: [new TextRun({ text: "Extreme price events (>200 €/MWh) cluster in 2021–2022; robust models should be evaluated with and without this crisis period.", size: 22 })]
      }),
      rule(),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 120 },
        children: [new TextRun({ text: "Leo Camberleng — EDHEC MSc Data Analysis & AI — May 2026", size: 18, italics: true, color: "888888" })]
      }),
    ]
  }],
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  }
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Report saved -> " + OUT);
});

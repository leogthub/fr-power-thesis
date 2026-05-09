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

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 22 })]
  });
}

function numItem(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22 })]
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
            new TextRun({ text: "Rapport de Modelisation -- Prix day-ahead francais", size: 18, color: "555555" }),
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
        children: [new TextRun({ text: "RAPPORT DE MODELISATION", size: 40, bold: true, color: "1F3864" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Prevision des Prix Day-Ahead de l'Electricite Francaise", size: 28, color: "2E5FA3" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "EDHEC MSc Data Analysis & AI  --  Memoire de Master", size: 20, italics: true, color: "777777" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 300 },
        children: [new TextRun({ text: "Periode de test : mai 2024 - avril 2025 | Sources : ENTSO-E, ERA5, ENTSOE-py, Yahoo Finance", size: 18, italics: true, color: "999999" })]
      }),
      rule(),

      // ── 1. METHODOLOGIE ──
      h1("1. Methodologie"),
      h2("1.1 Formulation du probleme"),
      p("L'objectif est de prevoir le prix day-ahead de l'electricite francaise (EPEX SPOT France) pour chaque heure du lendemain, en utilisant uniquement les informations disponibles avant la cloture du marche (~12h00 CET). Il s'agit d'un probleme de regression supervise sur serie temporelle a forte structure saisonniere, avec des prix variant de -100 a +3000 EUR/MWh selon les conditions de marche."),

      h2("1.2 Sources de donnees"),
      bullet("Prix et fondamentaux electriques : ENTSO-E Transparency Platform (2018-2025)"),
      bullet("Meteorologie : ERA5 reanalyse horaire (Copernicus CDS), France moyenne spatiale"),
      bullet("Prix combustibles : TTF gaz naturel, EUA carbone, charbon ARA (Yahoo Finance / Reuters)"),
      bullet("Flux transfrontaliers : echanges France-Allemagne, Espagne, Belgique, Royaume-Uni"),
      new Paragraph({ spacing: { before: 80, after: 100 }, children: [] }),

      h2("1.3 Feature Engineering"),
      p("35 features ont ete construites a partir des quatre sources de donnees :"),
      bullet("Calendrier : heure (sin/cos), jour de la semaine, mois (sin/cos), is_weekend"),
      bullet("Lags de prix : T-24h, T-48h, T-168h, moyennes glissantes 24h et 168h"),
      bullet("Fondamentaux : prevision de charge, generation nucleaire, gaz, hydro (fil+reservoir), solaire, eolien"),
      bullet("Flux transfrontaliers nets : FR-DE, FR-ES, FR-GB, FR-BE"),
      bullet("Prix combustibles : TTF (EUR/MWh), charbon ARA (EUR/t)"),
      bullet("Meteo ERA5 : temperature, vitesse vent, rayonnement solaire, precipitations"),
      bullet("Features derivees : HDD (seuil 17C - standard RTE France), CDD, proxy puissance eolienne"),
      bullet("Ratio disponibilite nucleaire : gen_nucleaire / 63 000 MW capacite installee (RTE 2023)"),
      bullet("Weather Stress Index : HDD / vitesse_vent (stress thermique x defaut eolien)"),
      bullet("Clean Spark Spread (CSS) et Clean Dark Spread (CDS)"),
      new Paragraph({ spacing: { before: 80, after: 100 }, children: [] }),

      h2("1.4 Split Train / Test"),
      p("Un split temporel strict est utilise pour eviter toute fuite de donnees (data leakage). Le modele est entraine sur 2018-2024 et evalue sur les 12 derniers mois (hors-echantillon strict)."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      makeTable([
        ["Ensemble", "Periode", "Heures"],
        ["Entrainement", "Janv. 2018 - Avr. 2024", "~55 200"],
        ["Test (hors-echantillon)", "Mai 2024 - Avr. 2025", "8 640"],
      ], [3500, 3500, 2026]),

      // ── 2. MODELES ──
      new Paragraph({ children: [new PageBreak()] }),
      h1("2. Modeles"),
      h2("Modele A -- Benchmark naif (lag-168h)"),
      p("Le modele naif predit le prix a l'heure h du jour D+1 egal au prix de l'heure h de la semaine precedente (D-7). Ce benchmark capture la saisonnalite hebdomadaire des prix, plus forte que la saisonnalite journaliere. Il constitue la barre minimale de performance que tout modele doit depasser."),

      h2("Modele B -- Random Forest sans meteo (ablation)"),
      p("Un Random Forest (500 arbres, profondeur max. 10, min. 5 obs. par feuille) est entraine sans aucune feature meteorologique. Ce modele d'ablation permet de mesurer la valeur ajoutee de la meteo par rapport aux seuls fondamentaux de marche et lags de prix."),

      h2("Modele C -- Random Forest avec meteo (modele principal)"),
      p("Le meme Random Forest entraine sur la matrice de features complete (35 variables incluant temperature, vent, HDD, WSI, etc.). Le Random Forest gere naturellement les interactions non-lineaires entre meteo et calendrier sans specification manuelle, est robuste aux outliers de prix, et fournit des importances de features interpretables."),

      h2("Modele D -- XGBoost avec meteo"),
      p("Un modele XGBoost (500 estimateurs, taux d'apprentissage 0.05, profondeur 6, subsample 0.8) entraine sur le meme jeu de features que le Modele C. XGBoost exploite un mecanisme de correction d'erreur sequentielle et une regularisation L1/L2 integree, ce qui le rend generalement performant sur les donnees tabulaires."),

      // ── 3. RESULTATS ──
      h1("3. Resultats"),
      h2("3.1 Performance sur le jeu de test"),
      p("Tous les modeles sont evalues sur les 8 640 heures de la periode de test (mai 2024 - avril 2025, entierement hors-echantillon). Les metriques retenues sont le MAE (Mean Absolute Error), RMSE (Root Mean Square Error), sMAPE (Symmetric MAPE, robuste aux prix negatifs), R2 et le Hit Ratio (precision directionnelle)."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      makeTable([
        ["Modele", "MAE (EUR/MWh)", "RMSE (EUR/MWh)", "sMAPE (%)", "R2", "Hit (%)"],
        ["A -- Naif lag-168h",        "33.09", "44.23", "70.9%", "0.160", "59.6%"],
        ["B -- RF sans meteo",        "16.95", "22.88", "44.2%", "0.775", "65.5%"],
        ["C -- RF avec meteo",        "16.94", "22.85", "44.2%", "0.776", "65.4%"],
        ["D -- XGBoost avec meteo",   "19.14", "28.18", "45.4%", "0.659", "65.8%"],
      ], [2800, 1600, 1700, 1500, 900, 1000]),
      new Paragraph({ spacing: { before: 120, after: 80 }, children: [] }),
      p("Les deux Random Forest (B et C) dominent largement le benchmark naif, reduisant le MAE de ~49% (33 vs 17 EUR/MWh) et ameliorant le R2 de 0.16 a 0.78. Fait notable : l'ajout de la meteo (C vs B) n'apporte qu'une amelioration marginale (+0.005 R2), suggere que les flux transfrontaliers, le prix du gaz TTF et la disponibilite nucleaire capturent deja la plupart de la variabilite liee aux conditions climatiques. XGBoost (D) sous-performe les Random Forests avec des hyperparametres standards, indiquant un potentiel de tuning."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      img("model_comparison.png", 560, 200) || p("[Figure : model_comparison.png]"),
      caption("Figure 1 -- MAE, RMSE, R2 et Hit Ratio des quatre modeles sur la periode de test (mai 2024 - avr. 2025)."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      h2("3.2 Tests de Diebold-Mariano (HLN 1997)"),
      p("Le test de Diebold-Mariano (1995) avec correction Harvey-Leybourne-Newbold (1997) evalue si les differences de performance entre modeles sont statistiquement significatives. H0 : accuracy predictive egale entre les deux modeles (MAE-base). Une statistique DM negative signifie que le modele 1 est plus precis."),
      new Paragraph({ spacing: { before: 100, after: 100 }, children: [] }),
      makeTable([
        ["Comparaison", "Stat. DM", "p-value", "Significativite"],
        ["C vs B : la meteo apporte-t-elle de la valeur ?", "-0.565", "0.572", "n.s."],
        ["C vs A : RF-meteo bat-il le naif ?",              "-53.53", "< 0.001", "***"],
        ["B vs A : RF sans meteo bat-il le naif ?",         "-53.44", "< 0.001", "***"],
        ["D vs C : XGBoost vs RF-meteo ?",                 "+10.88", "< 0.001", "***"],
      ], [3600, 1200, 1200, 1500]),
      new Paragraph({ spacing: { before: 120, after: 80 }, children: [] }),
      p("Interpretation : (1) Les deux RF battent le naif de maniere tres hautement significative (p < 0.001). (2) L'apport de la meteo n'est PAS statistiquement significatif (p = 0.57), resultat cle de l'etude d'ablation -- les features combustibles et flux transfrontaliers semblent absorber l'essentiel de l'information meteorologique. (3) XGBoost sous-performe significativement le RF avec meteo (DM = +10.88, p < 0.001), plaidant pour le RF comme modele de reference."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.3 Prevision vs Reel"),
      p("La figure 2 montre le premier mois de la periode de test. Le Modele C suit bien la structure intra-hebdomadaire et les pics journaliers. Les erreurs les plus importantes surviennent lors de baisses ou hausses soudaines liees a des evenements exogenes (sorties de centrales, chocs climatiques). Le benchmark naif (A) montre systematiquement un decalage sur les periodes de transition de marche."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("forecast_vs_actual.png", 560, 220) || p("[Figure : forecast_vs_actual.png]"),
      caption("Figure 2 -- Prevision vs reel, premier mois de test. Bleu = reel, rouge tirete = RF meteo (C), gris pointe = naif (A)."),

      // ── PAGE BREAK ──
      new Paragraph({ children: [new PageBreak()] }),

      h2("3.4 Distribution des erreurs"),
      p("Les erreurs de prevision sont approximativement centrees en zero, confirmant l'absence de biais systematique. Les queues de distribution sont plus lourdes pour les evenements extremes (prix >150 EUR/MWh), coherent avec la difficulte reconnue de prevoir les price spikes lies a des disruptions d'approvisionnement rares."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("error_distribution.png", 560, 200) || p("[Figure : error_distribution.png]"),
      caption("Figure 3 -- Distribution des erreurs (gauche) et dispersion reel vs predit (droite) pour les modeles B, C et D."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.5 Stabilite temporelle -- MAE mensuel"),
      p("La figure 4 decompose le MAE de test par mois pour les quatre modeles. Les erreurs sont plus elevees en automne-hiver (oct.-dec. 2024, janv. 2025) ou la volatilite des prix est maximale en raison des pics de demande de chauffage. Les deux RF affichent un profil mensuel tres similaire (ablation non concluante), tandis que XGBoost est plus volatile mois par mois."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("backtest_mae_by_month.png", 560, 185) || p("[Figure : backtest_mae_by_month.png]"),
      caption("Figure 4 -- MAE mensuel sur la periode de test (mai 2024 - avr. 2025) pour les quatre modeles."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      h2("3.6 Importance des variables -- Modele C"),
      p("La figure 5 montre le top 20 des features par importance MDI (Mean Decrease in Impurity) du Random Forest. Les lags de prix dominent (T-24h, T-168h, moyenne glissante 168h), confirmant la forte structure auto-regressive. Parmi les features meteo, la temperature (et HDD derive) et la vitesse du vent se classent hauts, validant l'hypothese centrale de la these. Le prix TTF du gaz et la disponibilite nucleaire apparaissent egalement en bonne position."),

      new Paragraph({ spacing: { before: 80 }, children: [] }),
      img("feature_importance_rf.png", 480, 310) || p("[Figure : feature_importance_rf.png]"),
      caption("Figure 5 -- Top 20 importances (Random Forest, Modele C). Rouge = meteo, bleu = autres."),

      // ── 4. DISCUSSION ──
      new Paragraph({ children: [new PageBreak()] }),
      h1("4. Discussion & Limites"),

      numItem("Prevision des price spikes : les deux modeles sous-estiment systematiquement les evenements extremes (>200 EUR/MWh). Une extension par classification de spike ou regression quantile pourrait ameliorer la performance en queue de distribution."),
      numItem("Ablation meteo non concluante : l'absence d'effet significatif de la meteo (test DM C vs B : p = 0.57) peut s'expliquer par la colinearite entre temperature et consommation de chauffage deja capturee par la prevision de charge ENTSO-E, et par le fait que le prix TTF du gaz integre deja les anticipations climatiques des marches."),
      numItem("Modele statique : l'entrainement est realise une seule fois sur 2018-2024. En production, un retrainement periodique (mensuel) serait necessaire pour s'adapter aux evolutions structurelles du marche (montee des EnR, evolution du mix nucleaire)."),
      numItem("Hyperparametres XGBoost : les hyperparametres utilises sont standards et non optimises par validation croisee temporelle. Un tuning par Bayesian optimization pourrait inverser le classement C vs D."),
      numItem("Horizon de prevision : seul l'horizon h+24h est traite. Une extension multi-horizon (h+1 a h+48) necessite soit une approche recursive (accumulation d'erreurs), soit un modele direct par horizon."),
      new Paragraph({ spacing: { before: 80, after: 100 }, children: [] }),

      rule(),
      h1("5. Conclusion"),
      p("Les deux modeles Random Forest (avec et sans meteo) surpassent significativement le benchmark naif sur la periode de test mai 2024 - avril 2025, atteignant un MAE ~16.9 EUR/MWh et R2 ~ 0.78 contre MAE 33.1 et R2 0.16 pour le naif. Le test de Diebold-Mariano confirme la superiorite statistique des RF (p < 0.001). L'apport marginal de la meteo n'est pas significatif dans ce jeu de features enrichi, ce qui constitue un resultat important pour la these : les fondamentaux de marche (prix combustibles, flux transfrontaliers, disponibilite nucleaire) semblent capturer l'essentiel de l'information meteo pertinente pour la prevision des prix day-ahead francais. Le Modele C (RF avec meteo) est retenu comme modele de reference pour sa completude interpretative et ses performances."),

      new Paragraph({ spacing: { before: 120 }, children: [] }),
      rule(),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 120 },
        children: [new TextRun({ text: "Leo Camberleng -- EDHEC MSc Data Analysis & AI -- Mai 2026", size: 18, italics: true, color: "888888" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Report saved -> " + OUT);
});

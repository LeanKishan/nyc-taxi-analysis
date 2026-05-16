// Generates the 10-page analytics report .docx using docx-js.
// Run: node reports/generate_report.js

const fs = require("fs");
const path = require("path");

// Force docx package to load from the global npm directory on Windows
const NPM_GLOBAL = path.join(
  process.env.APPDATA || "",
  "npm",
  "node_modules"
);
if (NPM_GLOBAL && !require("module").globalPaths.includes(NPM_GLOBAL)) {
  require("module").globalPaths.push(NPM_GLOBAL);
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition,
} = require(path.join(NPM_GLOBAL, "docx"));

// ---- Project paths ---------------------------------------------------------
const ROOT = "C:/Users/Kishan/nyc-taxi-analysis";
const FIG = path.join(ROOT, "reports", "figures");
const OUT = path.join(ROOT, "reports", "NYC_Taxi_2025_Analytics_Report.docx");

// ---- Figure metadata (true pixel sizes) -----------------------------------
const FIGURES = {
  daily_volume:        { w: 1560, h: 480 },
  demand_heatmap:      { w: 1560, h: 600 },
  fare_vs_distance:    { w: 1080, h: 720 },
  feature_importance:  { w: 960,  h: 600 },
  predicted_vs_actual: { w: 840,  h: 840 },
  speed_by_hour:       { w: 1320, h: 540 },
  tipping_behaviour:   { w: 1680, h: 540 },
  trip_distributions:  { w: 1800, h: 480 },
};

// Content width on US Letter w/ 1" margins is 6.5" = 624 px @ 96 DPI.
// We size each figure at most 540 px wide and let height follow the aspect.
function scaledFigure(name, captionText, opts = {}) {
  const meta = FIGURES[name];
  const maxW = opts.width || 540;
  const w = Math.min(maxW, meta.w);
  const h = Math.round(w * (meta.h / meta.w));
  const data = fs.readFileSync(path.join(FIG, `${name}.png`));
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 80 },
      children: [
        new ImageRun({
          type: "png",
          data,
          transformation: { width: w, height: h },
          altText: { title: name, description: captionText, name },
        }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [
        new TextRun({ text: `Figure: ${captionText}`, italics: true, size: 18, color: "555555" }),
      ],
    }),
  ];
}

// ---- Helpers for paragraphs ------------------------------------------------
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore: true,
    children: [new TextRun({ text })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text })],
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 140, line: 300 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, ...opts })],
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text })],
  });
}

// ---- KPI table for executive summary --------------------------------------
function kpiRow(label, value) {
  const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
  const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
  const lblWidth = 4680;
  const valWidth = 4680;
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: lblWidth, type: WidthType.DXA },
        shading: { fill: "F4F6F8", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true })] })],
      }),
      new TableCell({
        borders,
        width: { size: valWidth, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: value })] })],
      }),
    ],
  });
}

const kpiTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  rows: [
    kpiRow("Dataset", "NYC TLC Yellow Taxi (Jan–Mar 2025)"),
    kpiRow("Raw rows downloaded", "11,198,026"),
    kpiRow("Rows after cleaning", "8,516,174  (76.1%)"),
    kpiRow("Total trip revenue analyzed", "$236,025,351"),
    kpiRow("Average fare", "$18.15"),
    kpiRow("Average trip duration", "15.1 minutes"),
    kpiRow("Average trip distance", "3.22 miles"),
    kpiRow("ML model — validation MAE", "2.77 minutes"),
    kpiRow("ML model — validation R²", "0.864"),
  ],
});

// ---- Title page ------------------------------------------------------------
const titlePage = [
  new Paragraph({ spacing: { before: 2200, after: 200 }, children: [new TextRun({ text: "" })] }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
    children: [new TextRun({ text: "NYC YELLOW TAXI 2025", bold: true, size: 56, color: "1F3A5F" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 800 },
    children: [new TextRun({ text: "Analytics Report", size: 44, color: "1F3A5F" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text: "An end-to-end analysis of 8.5 million NYC taxi trips", italics: true, size: 26, color: "555555" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 1600 },
    children: [new TextRun({ text: "January – March 2025", size: 24, color: "555555" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [new TextRun({ text: "Prepared by", size: 22, color: "555555" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 600 },
    children: [new TextRun({ text: "Kishan", size: 28, bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [
      new TextRun({ text: "GitHub: ", size: 20, color: "555555" }),
      new ExternalHyperlink({
        children: [new TextRun({ text: "github.com/LeanKishan/nyc-taxi-analysis", style: "Hyperlink", size: 20 })],
        link: "https://github.com/LeanKishan/nyc-taxi-analysis",
      }),
    ],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 60 },
    children: [new TextRun({ text: "Data source: NYC Taxi & Limousine Commission", size: 20, color: "555555" })],
  }),
];

// ---- Table of contents page -----------------------------------------------
const tocPage = [
  new Paragraph({
    pageBreakBefore: true,
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: "Table of Contents" })],
  }),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
];

// ---- Section 1: Executive Summary -----------------------------------------
const section1 = [
  h1("1. Executive Summary"),
  p(
    "This report presents an end-to-end analysis of New York City Yellow Taxi trip records for the first quarter of 2025, covering more than 11 million raw trip records published by the NYC Taxi and Limousine Commission. After a rigorous cleaning pipeline, the working dataset contains 8.5 million trips representing roughly $236 million of fare revenue."
  ),
  p(
    "The analysis pursues four objectives: (1) characterize when and where demand for taxi service is highest; (2) quantify how traffic conditions affect trip economics; (3) compare airport service to ordinary city trips; and (4) build a predictive model that can estimate trip duration at the moment a trip begins."
  ),
  p(
    "Headline numbers are summarized below. Detailed methodology, exploratory findings, and modeling results follow in subsequent sections."
  ),
  kpiTable,
  h2("Top findings"),
  bullet("Demand follows a clear weekday-evening peak (5–6 PM) and a secondary late-night peak on Friday and Saturday."),
  bullet("Average vehicle speed drops from approximately 14 mph late at night to under 10 mph in the late afternoon — a quantifiable congestion signal."),
  bullet("Credit card payments dominate (86% of trips), with a median tip percentage clustered tightly around 20%."),
  bullet("Airport trips are far more expensive and longer (avg. $71 fare, 47 minutes) but tip a lower percentage (20% vs 26% for regular trips)."),
  bullet("A LightGBM regression model predicts trip duration with mean absolute error of 2.77 minutes and R² of 0.864 using only information available at trip start."),
];

// ---- Section 2: Dataset and Context ---------------------------------------
const section2 = [
  h1("2. Dataset and Context"),
  p(
    "The NYC Taxi and Limousine Commission (TLC) publishes monthly trip-record files in Apache Parquet format on a public CloudFront CDN. Each Yellow Taxi record describes a single completed trip and includes pickup and drop-off timestamps, taxi-zone identifiers, distance, fare components, payment type, and the trip's rate code (standard meter, JFK flat rate, Newark, negotiated, etc.)."
  ),
  p(
    "For this analysis we worked with three months of 2025 data — January, February, and March — totaling 189.5 MB of Parquet input and 11,198,026 trip records before cleaning. Yellow taxis were chosen over Green taxis and the FHV/FHVHV (Uber/Lyft) datasets because Yellow taxi service has the longest historical record and the most consistent schema."
  ),
  h2("Why this dataset"),
  p(
    "The TLC data is one of the most-cited open datasets in transportation analytics. It is granular enough to surface fine-grained patterns (hour-by-hour, zone-by-zone) yet large enough to require the same engineering discipline that any production data pipeline demands: streaming reads, efficient columnar formats, sanity filters, and an analysis-ready intermediate layer."
  ),
  h2("Key fields used"),
  bullet("pickup and dropoff datetime — power all temporal analyses and the trip-duration target."),
  bullet("trip_distance, fare_amount, total_amount — the trip economics layer."),
  bullet("PULocationID, DOLocationID — TLC taxi-zone codes (1–265) used for geographic analysis."),
  bullet("payment_type, tip_amount — drive the tipping behaviour analysis (tips are recorded only for card payments)."),
  bullet("RatecodeID — surfaces airport flat-rate trips, used as a feature in the ML model."),
];

// ---- Section 3: Methodology ----------------------------------------------
const section3 = [
  h1("3. Methodology"),
  h2("Pipeline overview"),
  p(
    "The project is structured as a five-stage pipeline. Each stage is implemented as a standalone Python module so that any stage can be re-run independently."
  ),
  bullet("Download — streams the monthly Parquet files from the TLC CDN with progress reporting and resume-safe behaviour."),
  bullet("Clean — applies sanity filters to drop trips with negative fares, zero or extreme durations, invalid zone IDs, or out-of-period timestamps."),
  bullet("Feature engineering — adds derived columns used by both the analytical reports and the ML model: hour of day, day of week, weekend flag, rush-hour flag, average speed, fare per mile, tip percentage, and airport-trip flag."),
  bullet("SQL analytics — DuckDB queries Parquet directly to produce eight business-question reports without a separate ETL step."),
  bullet("Modeling — a LightGBM regressor predicts trip duration in minutes from features known at trip start."),
  h2("Cleaning details"),
  p(
    "Approximately 23.9% of the raw rows were dropped during cleaning. The largest contributors were trips with implausibly short or long duration (under 1 minute or over 3 hours), trips with zero distance, and a small but real population of rows with timestamps outside the actual filing period — a quirk of the TLC pipeline where stale or misdated trips sometimes leak into the wrong monthly file."
  ),
  h2("Technology stack"),
  bullet("Python 3.12 — primary language."),
  bullet("pandas, pyarrow — in-memory data handling and Parquet I/O."),
  bullet("DuckDB — SQL analytics directly against Parquet files."),
  bullet("matplotlib, seaborn, plotly — static and interactive visualization."),
  bullet("scikit-learn, LightGBM — gradient-boosted regression model."),
  bullet("Streamlit — interactive web dashboard."),
  bullet("pytest — unit tests for the cleaning module."),
];

// ---- Section 4: Temporal Patterns -----------------------------------------
const section4 = [
  h1("4. Temporal Patterns"),
  p(
    "When are people taking taxis? The hour-of-day by day-of-week heatmap below answers that question with one image. The dark bands tell the story: weekday afternoons and evenings dominate, with a clear evening rush concentrated between 4 PM and 7 PM Tuesday through Friday. A secondary peak emerges late on Friday and Saturday nights, reflecting the city's nightlife economy. Sunday is consistently the quietest day of the week."
  ),
  ...scaledFigure("demand_heatmap", "Trips by hour of day × day of week"),
  p(
    "The daily volume series shows weekly seasonality and surfaces a handful of obvious dips. The largest drops correspond to severe-weather days and federal holidays. There is no large overall trend across the three months, which is consistent with NYC's relatively stable taxi market in early 2025."
  ),
  ...scaledFigure("daily_volume", "Daily trip volume across the analysis window"),
];

// ---- Section 5: Trip Economics --------------------------------------------
const section5 = [
  h1("5. Trip Economics"),
  p(
    "Three distributions characterize the bulk of NYC taxi activity: distance, duration, and fare. Each is heavily right-skewed — the median trip is short (1.5 miles, 10 minutes, $11) while a long tail of airport runs and outer-borough trips stretches the mean upward."
  ),
  ...scaledFigure("trip_distributions", "Distributions of trip distance, duration, and fare amount"),
  h2("Fare versus distance"),
  p(
    "Plotting fare against distance reveals a clean linear core — the meter rate — plus a distinct cluster around the $70 mark consistent with the JFK flat rate. This visualization, more than any single SQL query, justifies treating airport trips as a separate population in downstream analysis."
  ),
  ...scaledFigure("fare_vs_distance", "Fare amount versus trip distance"),
];

// ---- Section 6: Tipping ---------------------------------------------------
const section6 = [
  h1("6. Tipping Behaviour"),
  p(
    "Tips are only recorded for credit-card transactions — cash tips never reach the data feed. Among credit-card trips, the distribution is remarkably tight: the median tip percentage is approximately 20% and the bulk of the distribution sits in the 15–25% range. The clustering reflects the in-cab payment interface, which offers preset tip percentages and anchors riders to those defaults."
  ),
  ...scaledFigure("tipping_behaviour", "Tip percentage distribution and tip percentage by hour of day"),
  p(
    "Tip percentage is roughly constant across the day, with only modest variation. Riders are most generous in absolute dollar terms during off-peak nighttime hours, simply because fares are slightly higher and the percentage stays flat."
  ),
];

// ---- Section 7: Congestion ------------------------------------------------
const section7 = [
  h1("7. Congestion and Vehicle Speed"),
  p(
    "Trip duration alone does not separate long trips from slow trips. Computing the implied average speed (distance divided by duration) reveals the city's traffic pattern directly. Average speed dips from roughly 14 mph late at night to under 10 mph during the late-afternoon rush — a 30% reduction in throughput that any rider can feel."
  ),
  ...scaledFigure("speed_by_hour", "Average trip speed by hour of day"),
  p(
    "This is the analytical foundation for the ML model: the same trip distance produces a very different duration depending on the hour, so hour-of-day must be a feature."
  ),
];

// ---- Section 8: Machine Learning ------------------------------------------
const section8 = [
  h1("8. Trip-Duration Prediction Model"),
  h2("Problem framing"),
  p(
    "Estimating trip duration in advance is a foundational ride-hailing problem: dispatch must promise riders an ETA, surge-pricing logic depends on expected supply turnover, and downstream routing benefits from a duration estimate it can trust."
  ),
  p(
    "We frame the task as a regression problem with a hard constraint: only features available at the moment the trip starts may be used. This rules out any feature derived from the drop-off timestamp or final fare, preventing target leakage."
  ),
  h2("Features"),
  bullet("trip_distance and passenger_count — the basic trip request."),
  bullet("pickup_hour, pickup_dayofweek, pickup_month — temporal context."),
  bullet("pu_location_id, do_location_id — origin and destination zone (treated as categorical)."),
  bullet("ratecode_id — distinguishes standard meter from JFK / Newark / Nassau flat-rate trips."),
  bullet("Derived flags — is_weekend, is_rush_hour, is_airport_trip."),
  h2("Training and validation"),
  p(
    "A LightGBM gradient-boosted regressor was trained on 400,000 trips sampled from the cleaned dataset and validated on a 100,000-trip hold-out set. Early stopping based on validation RMSE prevented over-fitting; the best iteration was at 215 trees."
  ),
  p(
    "Validation performance: mean absolute error of 2.77 minutes, root-mean-square error of 4.49 minutes, coefficient of determination (R²) of 0.864. In practical terms, the typical prediction is within about three minutes of the true duration."
  ),
  ...scaledFigure("predicted_vs_actual", "Predicted versus actual trip duration on the validation set", { width: 400 }),
  ...scaledFigure("feature_importance", "LightGBM feature importance (gain)"),
  p(
    "Feature importance confirms the analytical findings: distance dominates, ratecode (which encodes the airport flat-rate trips) is the next strongest signal, and pickup hour captures most of the traffic effect. The model is interpretable in the sense that every meaningful feature corresponds to a real-world mechanism."
  ),
];

// ---- Section 9: Dashboard -------------------------------------------------
const section9 = [
  h1("9. Interactive Dashboard"),
  p(
    "All of the analytical findings in this report are also available as an interactive Streamlit dashboard. The dashboard is structured into five tabs:"
  ),
  bullet("Overview — top-level KPIs and a daily volume / revenue trend chart."),
  bullet("Time patterns — an interactive version of the hour × day-of-week demand heatmap."),
  bullet("Geography — the busiest pickup zones with average fare and total revenue."),
  bullet("Payments and tips — payment-type breakdown and the tip-percentage distribution."),
  bullet("ML predictor — a form that accepts trip parameters and calls the trained LightGBM model to produce a live duration estimate."),
  p(
    "Data caching with @st.cache_data and connection caching with @st.cache_resource keep navigation instantaneous after the first load. The dashboard is designed to deploy unchanged to Streamlit Community Cloud, with a bootstrap module that produces the analytical dataset on first launch."
  ),
];

// ---- Section 10: Recommendations and Future Work ------------------------
const section10 = [
  h1("10. Recommendations and Future Work"),
  h2("For the city"),
  bullet("Adjust signal timing around 5–7 PM weekday cycles in the highest-traffic zones — the 30% speed reduction during rush hour is a quantifiable inefficiency."),
  bullet("Investigate cash-only trips (12% of total) for service-quality gaps; they are concentrated in specific zones."),
  h2("For operators"),
  bullet("The airport flat-rate cluster suggests airport-dedicated dispatch could be modeled separately from city dispatch — the economics are sufficiently different."),
  bullet("The tight clustering of card tips at preset percentages is a clear UX result — small interface changes could lift overall tipping."),
  h2("For this project"),
  bullet("Geo visualization — join the TLC taxi-zone shapefile and produce a choropleth of trips, revenue, and average fare per zone using Folium or PyDeck."),
  bullet("Cross-modal comparison — pull the FHV (Uber / Lyft) data for the same period and compare market share, pricing, and demand patterns."),
  bullet("Probabilistic ETA — replace the point estimate with a quantile-regression model that gives P10–P90 intervals."),
  bullet("Deployment — already prepared. The dashboard can be hosted on Streamlit Community Cloud or Fly.io for a single-click public demo."),
  h2("Appendix: Reproducing the analysis"),
  p(
    "Every step in this report is reproducible from the public GitHub repository at github.com/LeanKishan/nyc-taxi-analysis. After cloning the repo and installing dependencies from requirements.txt, the entire pipeline runs with five commands: download, clean, run SQL analytics, train the model, and launch the dashboard. Unit tests for the cleaning module are included under tests/."
  ),
];

// ---- Build the document ---------------------------------------------------
const doc = new Document({
  creator: "Kishan",
  title: "NYC Yellow Taxi 2025 — Analytics Report",
  description: "End-to-end analytics report for NYC TLC Yellow Taxi Q1 2025 data",
  styles: {
    default: { document: { run: { font: "Calibri", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Calibri", color: "1F3A5F" },
        paragraph: { spacing: { before: 320, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: "1F3A5F" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "NYC Yellow Taxi 2025 — Analytics Report", italics: true, size: 18, color: "888888" })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "888888" }),
            new TextRun({ text: " of ", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: "888888" }),
          ],
        })],
      }),
    },
    children: [
      ...titlePage,
      ...tocPage,
      ...section1,
      ...section2,
      ...section3,
      ...section4,
      ...section5,
      ...section6,
      ...section7,
      ...section8,
      ...section9,
      ...section10,
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(OUT, buffer);
  const stat = fs.statSync(OUT);
  console.log(`Wrote ${OUT} (${(stat.size / 1024).toFixed(1)} KB)`);
});

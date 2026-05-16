// Generates the expanded analytics report .docx using docx-js.
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
  PageNumber,
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

// ---- Paragraph / heading helpers ------------------------------------------
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
function bulletBold(prefix, rest) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [
      new TextRun({ text: prefix, bold: true }),
      new TextRun({ text: rest }),
    ],
  });
}

// ---- Generic table helpers ------------------------------------------------
const CELL_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const BORDERS = { top: CELL_BORDER, bottom: CELL_BORDER, left: CELL_BORDER, right: CELL_BORDER };

function dataTable(headers, rows, columnWidths) {
  const total = columnWidths.reduce((a, b) => a + b, 0);
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((label, i) => new TableCell({
      borders: BORDERS,
      width: { size: columnWidths[i], type: WidthType.DXA },
      shading: { fill: "1F3A5F", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      children: [new Paragraph({
        children: [new TextRun({ text: label, bold: true, color: "FFFFFF" })],
      })],
    })),
  });
  const dataRows = rows.map((row, ri) => new TableRow({
    children: row.map((cellText, i) => new TableCell({
      borders: BORDERS,
      width: { size: columnWidths[i], type: WidthType.DXA },
      shading: { fill: ri % 2 === 0 ? "F4F6F8" : "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      children: [new Paragraph({
        alignment: i === 0 ? AlignmentType.LEFT : AlignmentType.RIGHT,
        children: [new TextRun({ text: String(cellText) })],
      })],
    })),
  }));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths,
    rows: [headerRow, ...dataRows],
  });
}

function tableCaption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 200 },
    children: [new TextRun({ text: `Table: ${text}`, italics: true, size: 18, color: "555555" })],
  });
}

// ---- KPI table for executive summary --------------------------------------
const kpiTable = dataTable(
  ["Metric", "Value"],
  [
    ["Dataset", "NYC TLC Yellow Taxi (Jan–Mar 2025)"],
    ["Raw rows downloaded", "11,198,026"],
    ["Rows after cleaning", "8,516,174  (76.1%)"],
    ["Raw data volume", "189.5 MB across 3 Parquet files"],
    ["Cleaned dataset size", "290.2 MB"],
    ["Total trip revenue analyzed", "$236,025,351"],
    ["Total tips collected (card)", "$30.3 million"],
    ["Average fare", "$18.15"],
    ["Average trip duration", "15.1 minutes"],
    ["Average trip distance", "3.22 miles"],
    ["Median tip percentage (card)", "20%"],
    ["ML model — validation MAE", "2.77 minutes"],
    ["ML model — validation R²", "0.864"],
    ["Lines of Python", "~900"],
    ["Unit tests", "5 (all passing)"],
  ],
  [4680, 4680]
);

// ---- Title page ------------------------------------------------------------
const titlePage = [
  new Paragraph({ spacing: { before: 2000, after: 200 }, children: [new TextRun({ text: "" })] }),
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
    spacing: { after: 200 },
    children: [new TextRun({ text: "January – March 2025", size: 24, color: "555555" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 1400 },
    children: [new TextRun({ text: "Data engineering · SQL analytics · Machine learning · Interactive dashboard", size: 20, color: "777777" })],
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
    "This report presents an end-to-end analytics project built on New York City Yellow Taxi trip records for the first quarter of 2025. The source data, published as monthly Apache Parquet files by the NYC Taxi and Limousine Commission (TLC), contains over 11 million raw trip records covering January through March 2025. After a multi-stage cleaning pipeline that removed sensor errors, implausible trips, and out-of-period timestamps, the analytical dataset contains 8,516,174 trips representing approximately $236 million in fare revenue and $30 million in tips."
  ),
  p(
    "The work is organized as a reproducible pipeline. A download module streams the monthly files from the TLC content delivery network. A cleaning and feature-engineering module applies sanity filters and constructs derived columns used throughout the analysis. A DuckDB-based analytics layer answers eight standing business questions directly against the Parquet output, with no intermediate ETL. A Jupyter notebook produces high-resolution figures suitable for both this report and for direct embedding in dashboards. A LightGBM gradient-boosted regression model predicts trip duration from features known at the moment a trip starts. Finally, a Streamlit dashboard with five interactive tabs exposes the cleaned data, the SQL reports, and the trained model to end users in a browser."
  ),
  p(
    "The analysis pursues five questions: when is demand for taxi service highest, how does congestion shape trip economics, how does payment behaviour vary across riders and times of day, how do airport trips differ from ordinary city trips, and to what extent is trip duration predictable from the information available at trip start. All five questions are answered quantitatively in this report and made interactive in the accompanying dashboard. Headline figures are summarized below; detailed methodology, exploratory findings, and modeling results follow in the subsequent sections."
  ),
  kpiTable,
  tableCaption("Headline metrics for the NYC Yellow Taxi 2025 project"),
  h2("Top findings at a glance"),
  bulletBold("Demand peaks at 5–6 PM on weekdays. ", "Hour 18 alone accounts for 621,501 trips, the single busiest hour. The second-highest hour is 17 with 611,553 trips. Together, 4 PM through 7 PM accounts for 26.5% of all trips despite being only 17% of the hours in a day."),
  bulletBold("Average speed drops by 30%+ during rush hour. ", "The average implied speed of a trip is 14.3 mph at midnight but falls to under 10 mph between 2 PM and 5 PM. This is the city's congestion footprint visible in the data."),
  bulletBold("Credit cards dominate (86%). ", "7.3 million card trips, 1.0 million cash trips, with disputed and zero-charge categories accounting for under 2%. Tips average 26% of fare on card trips; cash tips are not recorded in the data feed."),
  bulletBold("Airport runs are a separate economy. ", "287,809 airport trips averaged $71 in fare versus $16 for regular trips. Average distance was 17.9 miles versus 2.7 miles. Tip percentage was lower (20% vs 26%), but absolute tip dollars were higher."),
  bulletBold("Duration is predictable. ", "A LightGBM regressor achieves 2.77-minute mean absolute error on a 100,000-trip hold-out set, with R² of 0.864. Distance dominates the feature importance, followed by rate code and pickup hour."),
];

// ---- Section 2: Dataset and Context ---------------------------------------
const monthlyVolumeTable = dataTable(
  ["Month", "Raw rows", "Cleaned rows", "Retained", "File size"],
  [
    ["January 2025",  "3,475,226", "2,612,803", "75.2%", "59.2 MB"],
    ["February 2025", "3,610,485", "2,757,432", "76.4%", "60.3 MB"],
    ["March 2025",    "4,112,315", "3,145,939", "76.5%", "70.0 MB"],
    ["Total",         "11,198,026", "8,516,174", "76.1%", "189.5 MB"],
  ],
  [2200, 1850, 1850, 1730, 1730]
);

const section2 = [
  h1("2. Dataset and Context"),
  p(
    "The NYC Taxi and Limousine Commission has published anonymized trip records for the city's licensed for-hire vehicles since 2009. The Yellow Taxi dataset, which is the focus of this report, captures every completed metered or flat-rate trip taken in a yellow medallion taxi. The Green Boro Taxi, the For-Hire Vehicle (FHV) dataset, and the High-Volume FHV dataset (Uber, Lyft, Via, Juno) are published in parallel."
  ),
  p(
    "The 2025 monthly files are distributed as Apache Parquet, a columnar binary format that is both compact (the 3.5 million January records occupy 59.2 MB compressed) and amenable to direct SQL query without prior loading. Each record describes a single trip and includes the pickup and drop-off timestamps, the TLC taxi-zone identifiers for the pickup and drop-off locations (integer codes 1 through 265), trip distance in miles, the metered fare and a breakdown of surcharges and tolls, total amount paid, payment type, tip amount, passenger count as recorded by the driver, and the rate code that identifies whether the trip used the standard meter or one of several flat-rate categories (JFK, Newark, Nassau / Westchester, negotiated, group ride)."
  ),
  h2("Why Yellow Taxi specifically"),
  p(
    "Yellow Taxi was chosen for this analysis over the Green or FHV datasets for three reasons. First, Yellow Taxi has the longest continuous history of any TLC dataset, with consistent schema reaching back to 2009 — this makes it ideal for portfolio work where reproducibility against historical snapshots matters. Second, Yellow Taxis are restricted by regulation to certain pickup zones (predominantly Manhattan and the airports), which makes their data tractable for geographic analysis. Third, Yellow Taxi captures a meaningful portion of business travel and tourism, both of which produce interpretable behavioural signals — airport runs, business-district commutes, theatre district nightlife — that make findings easier to communicate."
  ),
  h2("Scope of this analysis"),
  p(
    "Three months of Yellow Taxi data — January, February, and March 2025 — were downloaded for this project. The choice of three months was deliberate: enough volume to surface stable patterns and to support a 500,000-trip ML training sample, but small enough that the entire pipeline runs end-to-end on a developer laptop in under 15 minutes."
  ),
  monthlyVolumeTable,
  tableCaption("Per-month row counts before and after cleaning"),
  p(
    "Monthly variation in volume reflects both calendar effects (March has 31 days versus 28 in February) and modest growth in baseline ridership over the quarter. Retention rates after cleaning are consistent across months, suggesting the cleaning filters are not biased toward any particular month."
  ),
  h2("Key fields used"),
  bulletBold("pickup_dt, dropoff_dt — ", "the two timestamps that power all temporal analyses and form the basis for the trip-duration target variable."),
  bulletBold("trip_distance, fare_amount, total_amount — ", "the trip-economics layer. trip_distance is the meter's distance reading in miles; fare_amount is the metered fare before tip and surcharges; total_amount is the full out-of-pocket cost to the rider."),
  bulletBold("pu_location_id, do_location_id — ", "TLC taxi-zone codes (1 to 265) used for geographic analysis. The TLC publishes a separate shapefile that maps each ID to a polygon on the street network."),
  bulletBold("payment_type, tip_amount — ", "drive the tipping behaviour analysis. Critically, tips are recorded only for credit card payments; cash tips are made directly from rider to driver and never reach the data feed."),
  bulletBold("ratecode_id — ", "identifies flat-rate trips. Codes 2 and 3 correspond to JFK and Newark airport runs respectively, and are used to construct the is_airport_trip feature in the ML model."),
];

// ---- Section 3: Methodology ----------------------------------------------
const techStackTable = dataTable(
  ["Layer", "Tool", "Purpose"],
  [
    ["Language", "Python 3.12", "All scripting and modeling"],
    ["Data I/O", "pyarrow + pandas", "Parquet reads, in-memory frames"],
    ["SQL analytics", "DuckDB 1.5", "Direct queries against Parquet files"],
    ["Visualization", "matplotlib + seaborn + plotly", "Static report figures and interactive charts"],
    ["Machine learning", "scikit-learn + LightGBM", "Train/validate split, gradient boosting"],
    ["Dashboard", "Streamlit", "Five-tab interactive web app"],
    ["Notebooks", "Jupyter", "Exploratory analysis"],
    ["Testing", "pytest", "Unit coverage for cleaning logic"],
    ["Version control", "Git + GitHub", "Public repository, deploy-ready"],
    ["Hosting (dashboard)", "Streamlit Cloud", "Free public deployment with GitHub integration"],
  ],
  [1800, 2280, 5280]
);

const cleanFiltersTable = dataTable(
  ["Field", "Accepted range", "Reason"],
  [
    ["trip_distance",       "0.1 – 100 mi",      "Reject zero and clearly invalid distances"],
    ["fare_amount",         "$0 – $500",         "Reject negatives (refunds) and extremes"],
    ["trip_duration_min",   "1 – 180 min",       "Reject zero-duration and >3-hour trips"],
    ["passenger_count",     "1 – 6",             "Standard cab capacity"],
    ["pu_location_id",      "1 – 265",           "Valid TLC zone range"],
    ["do_location_id",      "1 – 265",           "Valid TLC zone range"],
    ["total_amount",        "$0 – $1000",        "Reject negatives and clearly invalid totals"],
    ["pickup_dt",           "Within filing year","Reject stale or misdated rows"],
  ],
  [2200, 2500, 4660]
);

const section3 = [
  h1("3. Methodology"),
  p(
    "The project is structured as a five-stage data pipeline. Each stage is implemented as a standalone Python module under the src/ directory and can be re-run independently. The stages are: download, clean, feature-engineer, analyze (SQL), and model (ML). A separate dashboard layer in dashboard/ consumes the output of all five stages. A bootstrap module produces the analytical dataset on demand so the dashboard can be deployed to a cloud environment that does not have the source data."
  ),
  h2("Pipeline stages"),
  bulletBold("1. Download. ", "Streams the requested monthly Parquet files from the TLC CloudFront CDN. Skips files already on disk, retries on transient failures, and reports progress via tqdm."),
  bulletBold("2. Clean. ", "Reads the raw Parquet, normalizes column names across schema versions, computes trip_duration_min, and applies the sanity filters in the table below. Drops approximately 24% of raw rows."),
  bulletBold("3. Feature engineering. ", "Adds the derived columns used throughout the rest of the project: pickup_hour, pickup_dayofweek, pickup_day_name, pickup_month, pickup_date, is_weekend, is_rush_hour, is_airport_trip, time_of_day (categorical: Late night / Morning / Afternoon / Evening / Night), avg_speed_mph, fare_per_mile, tip_pct."),
  bulletBold("4. SQL analytics. ", "Eight named DuckDB queries that answer the standing business questions of the project. DuckDB reads directly from the cleaned Parquet — there is no intermediate database load step."),
  bulletBold("5. Modeling. ", "Trains a LightGBM gradient-boosted regressor on a 500,000-trip sample with an 80 / 20 train / validation split, early stopping on validation RMSE, and feature importance by gain."),
  h2("Cleaning filters"),
  p(
    "Each row in the raw Parquet is checked against the following ranges. A row is retained only if it passes every check. The ranges were chosen by inspecting the distribution of each field and identifying clear outlier regions."
  ),
  cleanFiltersTable,
  tableCaption("Sanity filters applied during cleaning"),
  p(
    "These filters drop 2,681,852 rows (23.9% of the raw input). The largest single contributor is implausibly short trips: rows with a recorded duration below one minute. The second largest is rows with a pickup timestamp from a different year — a recurring TLC data-quality issue where stale or misdated trips occasionally leak into the wrong monthly file. Without the year-range filter, time-series visualizations show a small but distracting cluster of rows from 2008, 2014, and other historical years."
  ),
  h2("Technology stack"),
  p(
    "The stack is deliberately conventional — no exotic dependencies, no proprietary services, no managed databases. Every component runs locally on a developer laptop and is reproducible from a fresh Python environment in under five minutes."
  ),
  techStackTable,
  tableCaption("Tools and their roles in the pipeline"),
];

// ---- Section 4: Temporal Patterns -----------------------------------------
const hourlyTable = dataTable(
  ["Hour", "Trips", "Avg fare", "Avg duration", "Avg speed"],
  [
    ["0",  "212,941", "$18.97", "13.4 min", "14.3 mph"],
    ["3",  " 61,209", "$16.94", "11.7 min", "14.8 mph"],
    ["5",  " 45,943", "$27.34", "16.9 min", "19.6 mph"],
    ["8",  "314,189", "$17.15", "15.0 min", "10.8 mph"],
    ["12", "465,702", "$17.93", "15.6 min", "10.0 mph"],
    ["15", "549,584", "$19.28", "17.4 min", " 9.7 mph"],
    ["17", "611,553", "$17.76", "15.9 min", " 9.7 mph"],
    ["18", "621,501", "$16.63", "14.3 min", "10.1 mph"],
    ["20", "496,981", "$17.56", "13.6 min", "12.2 mph"],
    ["23", "305,926", "$19.51", "14.1 min", "14.0 mph"],
  ],
  [1200, 2000, 1500, 2300, 2360]
);

const dowTable = dataTable(
  ["Day", "Trips", "Avg fare", "Avg tip %"],
  [
    ["Monday",    "1,050,599", "$19.27", "26.97%"],
    ["Tuesday",   "1,144,910", "$17.93", "26.31%"],
    ["Wednesday", "1,286,361", "$17.95", "26.21%"],
    ["Thursday",  "1,367,466", "$18.20", "26.10%"],
    ["Friday",    "1,303,764", "$18.12", "25.97%"],
    ["Saturday",  "1,319,501", "$16.99", "25.25%"],
    ["Sunday",    "1,043,596", "$18.92", "25.24%"],
  ],
  [2100, 2200, 2530, 2530]
);

const section4 = [
  h1("4. Temporal Patterns"),
  p(
    "The single most important variable in any analysis of urban movement is time. Demand for taxi service in New York City varies by an order of magnitude across a single day, and is reliably patterned by day of the week. The heatmap below presents the joint distribution of pickups by hour of day and day of week — one of the most informative single images in the entire dataset."
  ),
  ...scaledFigure("demand_heatmap", "Trips by hour of day × day of week"),
  p(
    "Several patterns are visible at a glance. Weekday afternoons through early evening (roughly 2 PM through 7 PM Tuesday, Wednesday, Thursday, and Friday) form a clear dark band that dominates the chart. A secondary, smaller dark band appears late Friday and Saturday nights, extending into the early morning hours of Saturday and Sunday — the city's nightlife economy at work. Sunday evening is conspicuously dimmer than the corresponding weekday evening windows, consistent with a population that is largely at home preparing for the work week."
  ),
  h2("Hour-of-day breakdown"),
  p(
    "Selected hours from the full 24-hour distribution are shown below to give a sense of how each metric varies across the day. The full table is available via the project's SQL analytics module."
  ),
  hourlyTable,
  tableCaption("Trip volume and economics across selected hours"),
  p(
    "Two observations from this table deserve attention. First, the busiest hour by trip volume — 6 PM (hour 18) with 621,501 trips — is not the hour with the highest average fare. Average fare peaks at the much quieter 5 AM hour ($27.34), driven by long-distance early morning runs to the airports for first-flight passengers. Second, average vehicle speed is at its lowest precisely during the busiest evening hours: 9.7 mph in both hour 15 and hour 17. The same trip from the same origin to the same destination is roughly 40% slower at 5 PM than at 3 AM."
  ),
  h2("Day-of-week breakdown"),
  p(
    "Aggregating across the full quarter, the day-of-week pattern follows a smooth Monday-to-Thursday ramp, with Thursday the single busiest day, before Friday and Saturday remain elevated and Sunday returns to weekday-low levels."
  ),
  dowTable,
  tableCaption("Trip volume and tipping by day of week"),
  p(
    "Average tip percentage declines slightly from Monday (27.0%) to Sunday (25.2%) — a modest but consistent effect. Several mechanisms could plausibly contribute: weekend trips tend to be shorter and more leisure-oriented (smaller absolute tip on a smaller fare may anchor to a smaller percentage), and weekend tipping is split across more cash transactions which never enter the percentage calculation."
  ),
  h2("Daily volume across the quarter"),
  p(
    "Plotting daily volume across the three months reveals weekly periodicity and a small number of anomalies. The biggest dips correspond to severe-weather days and federal holidays — Martin Luther King Jr. Day (January 20) and Presidents Day (February 17) both show visibly reduced volume. There is no strong overall trend across the three months, indicating that the underlying taxi market was stable across Q1 2025."
  ),
  ...scaledFigure("daily_volume", "Daily trip volume across the analysis window"),
];

// ---- Section 5: Trip Economics --------------------------------------------
const economicsTable = dataTable(
  ["Metric", "Median", "Mean", "P95"],
  [
    ["Distance (mi)",     "1.6",  "3.22",  "11.4"],
    ["Duration (min)",    "10.2", "15.1",  "39.6"],
    ["Fare ($)",          "$11.5","$18.15","$56.0"],
    ["Total amount ($)",  "$17.1","$27.71","$78.5"],
    ["Tip on card ($)",   "$2.6", "$4.13", "$12.0"],
    ["Fare per mile ($)", "$5.8", "$7.2",  "$15.4"],
  ],
  [3120, 2080, 2080, 2080]
);

const section5 = [
  h1("5. Trip Economics"),
  p(
    "Three distributions characterize the economic shape of NYC taxi activity: distance, duration, and fare. Each is heavily right-skewed — the median trip is short, but a long tail of airport runs and long outer-borough trips drags the mean well above the median. The visualization below presents the distributions side-by-side, each annotated with its median."
  ),
  ...scaledFigure("trip_distributions", "Distributions of trip distance, duration, and fare amount"),
  economicsTable,
  tableCaption("Summary statistics for the core trip metrics"),
  p(
    "The gap between the median and the mean is informative. A median distance of 1.6 miles versus a mean of 3.22 indicates that half of all trips are short — typical Manhattan-to-Manhattan rides — while the long tail of airport and outer-borough trips pulls the mean upward. The P95 column shows where the long tail sits: the 95th-percentile trip is 11.4 miles long, 39.6 minutes in duration, and costs $56 in fare alone. The very longest trips (the JFK and Newark airport runs) are the primary driver of this tail."
  ),
  h2("Fare versus distance"),
  p(
    "Plotting fare against distance shows the meter's pricing function with great clarity. The dense linear core represents standard metered trips, where the fare is approximately proportional to distance with a modest fixed component. A clearly distinct horizontal cluster at approximately $70 corresponds to the JFK flat rate (an $80 trip including the standard surcharge structure), which is independent of distance. Smaller clusters at other flat-rate price points are visible as well."
  ),
  ...scaledFigure("fare_vs_distance", "Fare amount versus trip distance"),
  p(
    "This visualization is itself a justification for treating airport trips as a separate population in downstream analysis. Their fares are bounded above by the flat rate regardless of distance, and the rate code (ratecode_id) cleanly distinguishes them. The ML model in section 8 leverages this separation explicitly."
  ),
  h2("Fare per mile"),
  p(
    "Dividing fare by distance gives a per-mile economic intensity. The mean across all trips is $7.20 per mile, but this varies systematically: short trips inside Manhattan have a much higher per-mile cost (because the fixed pickup fee amortizes over a small distance), while long airport runs have a much lower per-mile cost. Off-peak hours show a lower per-mile cost than rush hour because the per-minute waiting-time charge is less prominent when the cab is moving."
  ),
];

// ---- Section 6: Tipping ---------------------------------------------------
const paymentTable = dataTable(
  ["Payment type", "Trips", "Share", "Avg fare", "Avg tip", "Avg tip %"],
  [
    ["Credit card", "7,329,578", "86.07%", "$18.10", "$4.13", "26.0%"],
    ["Cash",        "1,042,774", "12.24%", "$18.08", "$0.00", "n/a"],
    ["Dispute",     "  109,568", " 1.29%", "$21.87", "$0.01", "n/a"],
    ["No charge",   "   34,277", " 0.40%", "$18.40", "$0.01", "n/a"],
  ],
  [2080, 1700, 1300, 1300, 1480, 1500]
);

const section6 = [
  h1("6. Payment and Tipping Behaviour"),
  p(
    "Payment behaviour in NYC taxis follows a clear hierarchy: credit cards dominate every other channel by a wide margin. The transition from cash-heavy to card-heavy payment was largely complete by 2015, but the residual cash share remains material — primarily for short, low-fare trips and in certain neighborhoods. The table below breaks down the four payment categories observed in the data."
  ),
  paymentTable,
  tableCaption("Trip volume and economics by payment type"),
  p(
    "Three points are worth highlighting. First, only 86% of trips have a recorded tip — the other 14% are cash, dispute, or no-charge categories. This is a critical data limitation: any tipping analysis is necessarily an analysis of card tipping. Second, average fare is remarkably stable across payment types — riders do not appear to use payment type to signal trip importance. Third, the small but meaningful dispute category (109,568 trips, 1.29%) has the highest average fare, suggesting that disputes are more likely to arise on more expensive trips where the absolute stakes are higher."
  ),
  h2("Card tipping distribution"),
  p(
    "Within the card-payment population, the distribution of tip percentage is remarkably tight. The median tip percentage is 20%, and the bulk of the distribution sits in the 15% to 25% range. The clustering reflects the in-cab payment interface, which presents the rider with preset percentage buttons (typically 15%, 20%, and 25%) and anchors most riders to those defaults."
  ),
  ...scaledFigure("tipping_behaviour", "Tip percentage distribution and tip percentage by hour of day"),
  p(
    "Plotting average tip percentage by hour of day reveals only modest variation: most hours sit between 25% and 27%, with a slight upward drift in late evening hours. This is consistent with the interpretation that the in-cab interface dominates tipping behaviour. Behavioural differences across time of day exist, but they are small relative to the structural anchoring effect of the preset buttons."
  ),
  h2("Implications"),
  p(
    "If a hypothetical operator wanted to lift overall tipping, the data suggest that small interface-level changes (adjusting the preset percentages, shifting the default selection) would have a far larger effect than any campaign aimed at changing rider behaviour. The clustering is too tight to be explained by rider deliberation."
  ),
];

// ---- Section 7: Congestion ------------------------------------------------
const rushTable = dataTable(
  ["Period", "Trips", "Avg speed", "Avg duration", "Fare / mile"],
  [
    ["Off-peak", "6,031,764", "11.6 mph", "15.0 min", "$7.83"],
    ["Rush hour", "2,484,410", " 9.5 mph", "16.4 min", "$8.91"],
  ],
  [1900, 1900, 1900, 1900, 1760]
);

const section7 = [
  h1("7. Congestion and Vehicle Speed"),
  p(
    "Duration alone does not separate long trips from slow trips: a 30-minute trip can be 25 miles on the highway or 4 miles inside Manhattan at 5 PM. Computing the implied average speed (distance divided by duration) gives a far more interpretable measure of congestion. Plotting average speed by hour of day produces the clearest single illustration of when New York's traffic is actually bad."
  ),
  ...scaledFigure("speed_by_hour", "Average trip speed by hour of day"),
  p(
    "The pattern is unambiguous. Late-night and early-morning hours show the highest speeds, peaking around 19.6 mph at 5 AM. Speed declines through the morning, reaches a plateau around 10 mph through the middle of the day, drops further into the late afternoon and early evening, and bottoms out at approximately 9.7 mph in the 3 PM to 5 PM window. The full swing from peak to trough is roughly 10 mph, or just over a 50% reduction in throughput."
  ),
  h2("Rush hour versus off-peak"),
  p(
    "Aggregating the hours into binary rush-hour (7–9 AM and 4–7 PM on weekdays) and off-peak categories quantifies the economic effect of congestion."
  ),
  rushTable,
  tableCaption("Rush hour versus off-peak — speed and pricing"),
  p(
    "Rush-hour trips are 18% slower on average (9.5 mph versus 11.6 mph) and cost 14% more per mile ($8.91 versus $7.83). The per-mile fare difference is the direct economic translation of slower traffic: New York's meter charges both by distance and by waiting time, so a slow trip costs more even for the same distance."
  ),
  h2("Why this matters for the ML model"),
  p(
    "The fact that the same origin-destination pair has a substantially different duration depending on hour of day is the central reason why the trip-duration prediction model in the next section must include temporal features. A model that uses only distance would systematically over-estimate duration off-peak and under-estimate it during rush hour. The pickup_hour feature, the is_rush_hour flag, and the location-pair categorical features together allow the model to learn this structure from data."
  ),
];

// ---- Section 8: Machine Learning ------------------------------------------
const hyperparamTable = dataTable(
  ["Hyperparameter", "Value", "Rationale"],
  [
    ["objective",        "regression",     "Continuous target (minutes)"],
    ["metric",           "rmse",           "Penalize large errors more strongly"],
    ["learning_rate",    "0.05",           "Moderate step size with many trees"],
    ["num_leaves",       "63",             "Allow moderate model capacity"],
    ["feature_fraction", "0.9",            "Modest random feature subsetting per tree"],
    ["bagging_fraction", "0.8",            "Row subsampling for variance reduction"],
    ["bagging_freq",     "5",              "Re-sample rows every 5 boosting rounds"],
    ["num_boost_round",  "500 (max)",      "Upper bound; early stopping stops sooner"],
    ["early_stopping",   "30 rounds",      "Stop when validation RMSE plateaus"],
  ],
  [2500, 2000, 4860]
);

const modelMetricsTable = dataTable(
  ["Metric", "Value", "Interpretation"],
  [
    ["MAE (validation)",  "2.77 min",        "Average absolute error in minutes"],
    ["RMSE (validation)", "4.49 min",        "Penalty for occasional large errors"],
    ["R² (validation)",   "0.864",           "Fraction of variance explained"],
    ["Best iteration",    "215",             "Boosting rounds before early stop"],
    ["Training rows",     "400,000",         "80% of 500K sample"],
    ["Validation rows",   "100,000",         "Held out, never seen during training"],
  ],
  [2400, 2000, 4960]
);

const section8 = [
  h1("8. Trip-Duration Prediction Model"),
  h2("Problem framing"),
  p(
    "Estimating trip duration in advance is a foundational ride-hailing problem. Dispatch systems must commit to an ETA when accepting a rider request, surge-pricing logic depends on expected supply turnover, and downstream routing algorithms benefit substantially from a duration estimate they can trust. The cost of being wrong is asymmetric: a substantially late arrival damages rider trust far more than an early arrival saves it."
  ),
  p(
    "We frame the task as a regression problem with a hard methodological constraint: only features available at the moment the trip starts may be used. This rules out any feature derived from the drop-off timestamp or the final fare, preventing target leakage. The constraint reflects how the model would actually be used — at trip-start, the dispatcher knows the origin, destination, distance estimate, and time, but not the future."
  ),
  h2("Features"),
  bulletBold("trip_distance — ", "the requested distance in miles. The single most important feature."),
  bulletBold("passenger_count — ", "weakly predictive of duration; large parties slightly increase boarding time."),
  bulletBold("pickup_hour, pickup_dayofweek, pickup_month — ", "temporal features that capture demand cycles and traffic conditions."),
  bulletBold("pu_location_id, do_location_id — ", "origin and destination zone IDs, treated as categorical. Allow the model to learn zone-pair-specific behaviour."),
  bulletBold("ratecode_id — ", "treated as categorical. Distinguishes standard metered trips from JFK / Newark / Nassau flat-rate trips."),
  bulletBold("is_weekend, is_rush_hour, is_airport_trip — ", "engineered boolean flags. The model can in principle infer these from the underlying fields, but providing them explicitly accelerates training."),
  h2("Training procedure"),
  p(
    "The cleaned 8.5-million-row dataset was sampled down to 500,000 trips with a fixed random seed to keep training tractable. The sample was split 80 / 20 into training (400,000 rows) and validation (100,000 rows) sets, again with a fixed seed for reproducibility. LightGBM was used as the gradient-boosting framework for its speed on tabular data and its native support for categorical features. The hyperparameter table below summarizes the configuration."
  ),
  hyperparamTable,
  tableCaption("LightGBM hyperparameters"),
  p(
    "Early stopping was triggered at 215 boosting rounds out of a possible 500, indicating that the model had reached its capacity for the available data. Training took approximately 30 seconds on a developer laptop. The categorical features (pu_location_id, do_location_id, ratecode_id) were declared explicitly so that LightGBM uses its optimal-split algorithm for categoricals rather than treating them as ordinal numbers."
  ),
  h2("Validation results"),
  modelMetricsTable,
  tableCaption("Validation metrics on the 100,000-trip hold-out set"),
  p(
    "The headline number is mean absolute error of 2.77 minutes. In plain language, the typical prediction is within about three minutes of the true duration. For a model that uses only trip-start information, this is a strong result. The R² of 0.864 indicates that the features collectively explain 86% of the variance in trip duration."
  ),
  ...scaledFigure("predicted_vs_actual", "Predicted versus actual trip duration on the validation set", { width: 400 }),
  p(
    "The predicted-versus-actual scatter plot above shows the model tracks the truth closely along the diagonal across the full range of durations, with predictable widening of the error band at very long durations (where there are simply fewer training examples)."
  ),
  ...scaledFigure("feature_importance", "LightGBM feature importance (gain)"),
  p(
    "Feature importance by gain ranks distance as the dominant feature by an order of magnitude. The second-most-important feature is the rate code, which encodes the flat-rate airport runs that the model treats as a structurally different population. Pickup hour comes third, capturing most of the traffic effect. The destination and origin zone IDs follow, allowing the model to learn that certain trips between specific zones are reliably faster or slower than the distance-and-hour baseline would suggest."
  ),
  h2("Limitations and possible extensions"),
  p(
    "Two limitations of the current model are worth flagging. First, the model produces a point estimate, not a probability distribution. For real dispatch use, a probabilistic ETA (for example, a P10 to P90 range) would be more useful. This could be obtained with a quantile-regression LightGBM ensemble. Second, the model does not include weather or real-time traffic features. Incorporating an hourly precipitation feed or a real-time traffic API would likely lift R² further, particularly during rare but impactful weather events."
  ),
];

// ---- Section 9: Dashboard -------------------------------------------------
const tabsTable = dataTable(
  ["Tab", "Purpose", "Key components"],
  [
    ["Overview",        "Top-level KPIs",            "Five KPI cards, daily volume / revenue chart, 14-day data table"],
    ["Time patterns",   "Hour-of-day demand",        "Interactive heatmap, hourly bar chart"],
    ["Geography",       "Pickup zones",              "Horizontal bar chart of top N zones, configurable via slider"],
    ["Payments & tips", "Payment breakdown",         "Pie chart, table, tip-percentage histogram with median line"],
    ["ML predictor",    "Live duration estimation",  "Input form with 9 fields, button-triggered prediction, model metadata"],
  ],
  [1700, 2700, 4960]
);

const section9 = [
  h1("9. Interactive Dashboard"),
  p(
    "All of the analytical findings in this report are exposed through a Streamlit dashboard that runs in a browser. The dashboard is structured into five tabs, each with a distinct analytical focus. Tabs are summarized in the table below."
  ),
  tabsTable,
  tableCaption("Streamlit dashboard structure"),
  h2("Performance and caching"),
  p(
    "All dataset loading is wrapped with Streamlit's caching decorators. The DuckDB connection is created once per session with @st.cache_resource so that subsequent tab navigation does not pay any setup cost. Each SQL query result is wrapped with @st.cache_data so that re-rendering a tab after a slider change reuses the previous result. The net effect is that the first page load takes one to two seconds; every subsequent interaction is effectively instantaneous."
  ),
  h2("Live ML inference"),
  p(
    "The ML predictor tab loads the trained LightGBM model at session start using @st.cache_resource (the model is approximately 1.5 MB on disk and loads in milliseconds). When the user submits the form, the inputs are assembled into a single-row pandas DataFrame, the categorical columns are converted to the correct dtype, and the model produces a duration prediction in under 10 milliseconds. The prediction is displayed alongside the implied average speed and a small contextual note (for example, a warning that rush-hour predictions have higher variance)."
  ),
  h2("Cloud deployment"),
  p(
    "The dashboard is designed to deploy to Streamlit Community Cloud unchanged. A bootstrap module included in src/ checks at app startup whether the cleaned dataset exists; if not, it downloads a single month of raw data and produces a 500,000-row analytical sample using DuckDB. This pipeline runs entirely within DuckDB rather than pandas to stay within the 1 GB RAM ceiling of the free deployment tier. The bootstrap typically completes in under fifteen seconds on first launch, after which the dataset is permanent for the lifetime of the deployed container."
  ),
];

// ---- Section 10: Recommendations and Future Work ------------------------
const section10 = [
  h1("10. Recommendations and Future Work"),
  h2("Recommendations for the city and operators"),
  p(
    "The single most actionable finding for transportation planners is the 50% range in average vehicle speed across the day. Signal timing optimization, congestion pricing during the worst-hit hours, and selective taxi-stand placement in the highest-volume zones are all interventions that could plausibly recover throughput. The data here can serve as a quantitative baseline for any such intervention."
  ),
  p(
    "For taxi operators, the airport flat-rate cluster suggests that airport-dedicated dispatch should be modeled separately from city dispatch. The economics — and the duration distributions — are sufficiently different that a single fleet-wide model is leaving signal on the table. Operators should also consider that the tight clustering of card tips around the in-cab preset percentages is a UX outcome more than a behavioural one: even small changes to the preset values could plausibly produce meaningful aggregate revenue changes."
  ),
  h2("Future work on this project"),
  bulletBold("Geographic visualization. ", "Join the TLC taxi-zone shapefile and render choropleths of trips, revenue, and average fare per zone using Folium or PyDeck. This would close a clear gap in the current analysis, which represents zones only as integer IDs."),
  bulletBold("Cross-modal comparison. ", "Pull the FHV (Uber, Lyft) data for the same period and compare market share, pricing, and demand patterns. Yellow Taxi is now a minority of the for-hire vehicle market in NYC; understanding the relative dynamics is important context."),
  bulletBold("Probabilistic ETA. ", "Replace the point-estimate trip-duration model with a quantile-regression ensemble that produces P10–P90 intervals. This is more useful in production and is a small extension of the current LightGBM training pipeline."),
  bulletBold("Weather features. ", "Join an hourly precipitation feed (NOAA publishes this freely for JFK and LGA) and incorporate weather conditions as additional model features. Rain and snow are known to substantially affect both demand and trip duration."),
  bulletBold("Real-time deployment. ", "Containerize the dashboard with Docker and deploy to Fly.io or Render for a permanent public demo, with auto-redeploy on every Git push."),
  bulletBold("Cohort analysis. ", "Pull a full calendar year and analyze seasonality at the year scale, including holiday effects, summer-versus-winter tourism patterns, and the effect of major events at Madison Square Garden, the United Nations, and other major venues."),
  h2("Lessons learned"),
  p(
    "Three lessons from this project are worth recording. First, public datasets are rarely as clean as their documentation suggests — the year-range filter alone caught hundreds of thousands of rows that would have polluted time-series visualizations in subtle ways. Always inspect the actual data distribution before trusting the schema. Second, DuckDB is dramatically more efficient than pandas for large-Parquet analytics in memory-constrained environments — the Streamlit Cloud bootstrap module was rewritten from pandas to DuckDB after the original deployment OOM'd on a 1 GB free-tier container, and the DuckDB version not only fits in memory but is also faster. Third, modeling decisions and analytical decisions should be made together: knowing that airport trips form a structurally different population shaped both the EDA (treating them as a separate cohort in every table) and the ML model (including ratecode_id as a categorical feature)."
  ),
  h2("Reproducing this analysis"),
  p(
    "Every step in this report is reproducible from the public GitHub repository at github.com/LeanKishan/nyc-taxi-analysis. After cloning the repository and installing dependencies from requirements.txt, the entire pipeline runs with five commands:"
  ),
  bullet("python -m src.download_data --year 2025 --months 1 2 3"),
  bullet("python -m src.data_cleaning"),
  bullet("python -m src.sql_analytics"),
  bullet("python -m src.train_model"),
  bullet("streamlit run dashboard/app.py"),
  p(
    "Unit tests for the cleaning module are under tests/ and can be run with pytest. The expected end-to-end runtime on a developer laptop is approximately fifteen minutes, dominated by the initial Parquet download and the ML training step."
  ),
  h2("Acknowledgments"),
  p(
    "Data is courtesy of the New York City Taxi and Limousine Commission, which publishes the trip records as a public service. The open-source ecosystem that makes this kind of analysis possible — DuckDB, LightGBM, pandas, Streamlit, scikit-learn — represents extraordinary value freely contributed by their respective maintainers and communities."
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

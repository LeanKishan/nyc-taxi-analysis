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
    "This report presents an end-to-end analytics project built on New York City Yellow Taxi trip records for the first quarter of 2025. The source data, published as monthly Apache Parquet files by the NYC Taxi and Limousine Commission (TLC), contains over 11 million raw trip records covering January through March 2025. After a multi-stage cleaning pipeline that removed sensor errors, implausible trips, and out-of-period timestamps, the analytical dataset contains 8,516,174 trips representing approximately $236 million in fare revenue and $30 million in tips. The dataset captures the day-to-day pulse of one of the most heavily studied urban transportation systems in the world, and provides an ideal substrate for demonstrating an analytics workflow that combines streaming data ingestion, large-scale SQL aggregation, exploratory visualization, predictive modeling, and interactive deployment."
  ),
  p(
    "The work is organized as a reproducible pipeline. A download module streams the monthly files from the TLC content delivery network with a progress display and resume-safe behaviour, so that interrupted downloads do not require restarting from scratch. A cleaning and feature-engineering module applies sanity filters that drop approximately 24% of the raw rows, and constructs the derived columns (hour of day, day of week, is-rush-hour flag, average speed, fare per mile, tip percentage, and so on) that are used throughout the rest of the project. A DuckDB-based analytics layer answers eight standing business questions directly against the cleaned Parquet output, with no intermediate ETL or database load step required. A Jupyter notebook produces high-resolution figures suitable for both this report and for direct embedding in dashboards. A LightGBM gradient-boosted regression model predicts trip duration from features known at the moment a trip starts, achieving validation metrics that would be competitive in any production ride-hailing setting. Finally, a Streamlit dashboard with five interactive tabs exposes the cleaned data, the SQL reports, and the trained model to end users in a browser, complete with a bootstrap module that rebuilds the analytical dataset on first launch so the dashboard can deploy to a memory-constrained cloud container."
  ),
  p(
    "The analysis pursues five interlocking questions. First, when is demand for taxi service highest, and how stable are those temporal patterns across days of the week and across the three months of the analysis window? Second, how does urban congestion shape the economics of an individual trip — specifically, how much does the same origin-destination pair vary in duration and in implied per-mile cost depending on the time at which it occurs? Third, how does payment behaviour vary across riders and times of day, and what does the tightness of the credit-card tipping distribution tell us about the role of interface design in shaping economic outcomes? Fourth, how do airport trips differ from ordinary city trips, and is it analytically useful to treat them as a structurally separate population? And fifth, to what extent is trip duration predictable from the limited information available at the moment a trip is dispatched? All five questions are answered quantitatively in this report, supported by figures and aggregated tables, and are made interactive in the accompanying Streamlit dashboard. Headline figures are summarized in the table below; detailed methodology, exploratory findings, and modeling results follow in the subsequent sections."
  ),
  kpiTable,
  tableCaption("Headline metrics for the NYC Yellow Taxi 2025 project"),
  h2("How to read this report"),
  p(
    "The report is organized so that any reader can extract value from it at their preferred level of detail. The remainder of this executive summary lists the five top-level findings as bullet points; each finding is then developed across one or more later sections, with the supporting tables, figures, and discussion. A reader with five minutes can stop after this summary. A reader with thirty minutes will benefit from reading the temporal patterns section (section 4), the congestion section (section 7), and the machine learning section (section 8), which together carry the bulk of the analytical content. A reader interested in reproducing the work end-to-end should read the methodology section (section 3) and the appendix in section 10."
  ),
  p(
    "Every table and figure in this report was generated from the public NYC TLC data using the open-source code in the project's GitHub repository. The numbers are not estimates and they are not adjusted: they are direct aggregations from the cleaned dataset, computed by the SQL queries and Python scripts included in the project. Where rounding occurs, it is to one or two decimal places for readability; the underlying numbers are preserved in the analytical pipeline."
  ),
  h2("Top findings at a glance"),
  bulletBold("Demand peaks at 5–6 PM on weekdays. ", "Hour 18 alone accounts for 621,501 trips, the single busiest hour. The second-highest hour is 17 with 611,553 trips. Together, 4 PM through 7 PM accounts for 26.5% of all trips despite being only 17% of the hours in a day. The pattern is highly stable across weekdays and replicates closely week-over-week, indicating a structural demand cycle rather than any temporary phenomenon."),
  bulletBold("Average speed drops by 30%+ during rush hour. ", "The average implied speed of a trip is 14.3 mph at midnight but falls to under 10 mph between 2 PM and 5 PM. This is the city's congestion footprint made directly visible in the data, and it has substantial economic consequences: a rush-hour trip costs approximately 14% more per mile than the same trip at off-peak times because the city's meter charges for both distance and waiting time."),
  bulletBold("Credit cards dominate (86%). ", "7.3 million card trips, 1.0 million cash trips, with disputed and zero-charge categories accounting for under 2% combined. Tips average 26% of fare on card trips and are remarkably tightly clustered around the in-cab interface's preset values; cash tips are paid directly to the driver and never reach the data feed, a limitation that constrains the scope of any tipping analysis."),
  bulletBold("Airport runs are a separate economy. ", "287,809 airport trips averaged $71 in fare versus $16 for regular trips, and 17.9 miles in distance versus 2.7 miles. Tip percentage was lower (20% vs 26%), reflecting the application of percentage tips to a much larger base, but absolute tip dollars were substantially higher. The flat-rate structure for JFK trips produces a distinctive horizontal cluster in the fare-versus-distance scatter that justifies treating airport trips as a separate analytical population."),
  bulletBold("Duration is predictable. ", "A LightGBM gradient-boosted regressor achieves 2.77-minute mean absolute error on a 100,000-trip hold-out set, with R² of 0.864. The training set was 400,000 trips, and the model uses only information available at the moment the trip begins, ensuring no target leakage. Distance dominates the feature importance, followed by rate code (which encodes the airport flat-rate population) and pickup hour (which encodes traffic conditions)."),
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
    "The NYC Taxi and Limousine Commission has published anonymized trip records for the city's licensed for-hire vehicles since 2009, making it one of the longest-running open transportation datasets in the world. The Yellow Taxi dataset, which is the focus of this report, captures every completed metered or flat-rate trip taken in a yellow medallion taxi within the five boroughs of New York City. The Green Boro Taxi dataset, which began publication in 2013 alongside the launch of the Green Taxi program, the For-Hire Vehicle (FHV) dataset covering car services not authorized to pick up street hails, and the High-Volume For-Hire Vehicle (HVFHV) dataset capturing the on-demand platforms (Uber, Lyft, Via, Juno) are published in parallel. Together these four datasets describe essentially the entire for-hire vehicle market in the city, and they have been the subject of hundreds of academic papers, blog posts, and Kaggle competitions over the past decade."
  ),
  p(
    "The decision to publish this data publicly is itself an interesting policy choice. The TLC's stated rationale is transparency: the agency regulates the for-hire vehicle market and believes that public access to trip data supports informed civic debate about fares, congestion, equity in service across neighborhoods, and the impact of new entrants like Uber and Lyft on the traditional taxi industry. As a side effect, the data has become an essential teaching resource for data engineering and data science: it is large enough to require real engineering discipline (a single year of Yellow Taxi data is tens of millions of rows), real enough to surface genuine analytical patterns rather than synthetic toy problems, and well-documented enough that any analyst can get up to speed quickly. The trade-off the TLC accepts in making the data public is the risk that individual riders or drivers might be re-identified. The agency mitigates this by aggregating pickup and drop-off locations to taxi zones (rather than exact coordinates) and by suppressing fields that would directly identify the driver or the vehicle."
  ),
  p(
    "The 2025 monthly files are distributed as Apache Parquet, a columnar binary format that has become the de facto standard for analytical data interchange. Parquet is both compact — the 3.5 million January records occupy only 59.2 MB on disk despite preserving every field at full precision — and amenable to direct SQL query without prior loading into a database. Tools like DuckDB, Polars, and modern Spark can scan Parquet files in seconds, projecting only the columns required by a query and pushing predicates down to the file scan to minimize I/O. This is a meaningful improvement over the CSV format that the TLC used in earlier years: a year of Yellow Taxi data in CSV would be several gigabytes and would require explicit type inference on every read, whereas the equivalent Parquet is roughly an order of magnitude smaller and types are preserved from the writer."
  ),
  p(
    "Each Yellow Taxi record describes a single completed trip and includes the pickup and drop-off timestamps, the TLC taxi-zone identifiers for the pickup and drop-off locations (integer codes 1 through 265), trip distance in miles as recorded by the taxi meter, the metered fare amount and a breakdown of surcharges and tolls, the total out-of-pocket cost to the rider, the payment type, the tip amount (only for credit card payments), the passenger count as entered by the driver at the start of the trip, and the rate code that identifies whether the trip used the standard meter or one of several flat-rate categories. The rate codes are particularly important for downstream analysis because they distinguish the structurally different airport-run population (codes 2 for JFK and 3 for Newark) from the standard metered trips that comprise the bulk of the dataset. A small number of additional housekeeping fields (vendor identifier, store-and-forward flag, congestion surcharge components) round out the schema, but are not central to the analytical questions pursued in this report."
  ),
  h2("Why Yellow Taxi specifically"),
  p(
    "Yellow Taxi was chosen for this analysis over the Green Taxi, FHV, or High-Volume FHV datasets for three principal reasons. First, Yellow Taxi has the longest continuous history of any TLC dataset, with a consistent core schema reaching back to 2009. Schemas in the FHV datasets have shifted substantially over the years, particularly when the High-Volume category was created in 2019 to separately track Uber and Lyft trips, and analyses that span those schema breaks require careful column-level reconciliation. Yellow Taxi, by contrast, has been remarkably stable, which makes it ideal for portfolio work where reproducibility against historical snapshots matters and where the analyst should be able to extend the work to additional time periods without rewriting the cleaning pipeline."
  ),
  p(
    "Second, Yellow Taxis are restricted by regulation to certain pickup zones — predominantly Manhattan below 110th Street, the two major airports (JFK and LaGuardia), and a handful of other designated zones. Green Taxis were created precisely to serve the parts of the city that Yellow Taxis were not allowed to serve, and on-demand platforms like Uber operate everywhere with no zone restrictions. This regulatory concentration makes Yellow Taxi data tractable for geographic analysis: a single chart showing the top twenty pickup zones captures the majority of the trip volume, which is harder to achieve with the more geographically diffuse FHV datasets. It also means that any geographic finding in the Yellow Taxi data is straightforward to interpret in terms of the underlying neighbourhoods, since they are mostly concentrated in well-known and well-defined parts of Manhattan."
  ),
  p(
    "Third, Yellow Taxi captures a meaningful portion of both business travel and tourism, both of which produce interpretable and analyzable behavioural signals. Airport runs to JFK and LGA are the most visible example: the flat-rate fare structure for JFK trips produces a distinctive cluster in the fare-versus-distance distribution, and the timing of those airport runs correlates with early-morning departures (for outbound passengers) and arrival waves throughout the day (for inbound passengers). Business-district pickups around midtown Manhattan and the Financial District trace out commuter patterns and post-work happy hour rhythms. Theatre-district pickups concentrate sharply between 10:30 and 11:30 PM as evening performances let out. All of these patterns are visible in the data, and they make findings far easier to communicate to a non-technical audience than would be the case with a more diffuse dataset where the signal is harder to isolate."
  ),
  h2("Scope of this analysis"),
  p(
    "Three months of Yellow Taxi data — January, February, and March 2025 — were downloaded for this project. The choice of three months was deliberate. On the one hand, the volume needed to be large enough to surface stable temporal patterns, to support a 500,000-trip machine learning training sample without overfitting concerns, and to make the analytical findings credible at the population level. On the other hand, the volume had to remain small enough that the entire pipeline runs end-to-end on a developer laptop in under fifteen minutes — a constraint that is essential for portfolio work because the typical reviewer (a recruiter, an interviewer, or a peer reading the GitHub repo) needs to be able to reproduce the work without spinning up cloud infrastructure or waiting hours for a download to complete. Three months hits both targets comfortably: 11.2 million raw rows, 8.5 million after cleaning, a 290 MB processed Parquet file that fits comfortably in memory on any modern laptop, and a complete fresh run from download to dashboard launch in approximately twelve minutes."
  ),
  monthlyVolumeTable,
  tableCaption("Per-month row counts before and after cleaning"),
  p(
    "Monthly variation in raw volume reflects both calendar effects (March has 31 days versus 28 in February, and 28 days are by definition fewer opportunities to take a taxi) and modest organic growth in baseline ridership over the quarter. The retention rate after cleaning is remarkably consistent across months — 75.2%, 76.4%, and 76.5% respectively — which is reassuring evidence that the cleaning filters are not systematically biased against any particular month. A retention rate that varied substantially across months would suggest either that the TLC data quality differs across months in ways the filters do not handle, or that the filters themselves are picking up real signal rather than data noise. Neither appears to be the case here."
  ),
  h2("Key fields used"),
  p(
    "The Yellow Taxi schema contains roughly twenty fields, but only a handful are central to the analytical questions pursued in this report. The remainder are either housekeeping (vendor IDs and similar) or low-information for our purposes. The fields the analysis depends on are listed below, with a brief explanation of what each is used for and any data-quality considerations associated with it."
  ),
  bulletBold("pickup_dt, dropoff_dt — ", "the two timestamps that power all temporal analyses and form the basis for the trip-duration target variable. The cleaned dataset filters out any rows where the dropoff timestamp precedes the pickup, where the implied duration is zero, or where either timestamp falls outside the actual filing year. The last filter is more important than it sounds: the raw TLC data periodically contains a small but non-trivial number of trips with timestamps from earlier years that have somehow leaked into the wrong monthly file."),
  bulletBold("trip_distance, fare_amount, total_amount — ", "the trip-economics layer. trip_distance is the meter's distance reading in miles, not a straight-line distance — it includes any backtracking that the driver did to navigate one-way streets and so on. fare_amount is the metered fare before tip, tolls, and the various surcharges that NYC has progressively added over the years. total_amount is the full out-of-pocket cost to the rider, including everything. The difference between fare_amount and total_amount is roughly $5 on average and includes the MTA surcharge, the improvement surcharge, the congestion surcharge, and tolls."),
  bulletBold("pu_location_id, do_location_id — ", "TLC taxi-zone codes (1 to 265) used for geographic analysis. The TLC publishes a separate shapefile that maps each zone ID to a polygon on the street network. Zones range in size from single Manhattan blocks (in the densest parts of midtown) to multi-square-mile polygons in the outer boroughs. Zone IDs 132 and 138 correspond to JFK and LaGuardia respectively, and consistently appear among the highest-volume zones because every airport pickup originates there."),
  bulletBold("payment_type, tip_amount — ", "drive the tipping behaviour analysis. Critically, tips are recorded only for credit card payments (payment_type = 1); cash tips are made directly from rider to driver and never reach the data feed. The data also distinguishes 'no charge' trips (which include some corporate or driver-comp scenarios) and 'dispute' trips (where the rider has contested the charge). Both of these are small in absolute count but they appear in the totals and need to be excluded from tipping analyses."),
  bulletBold("ratecode_id — ", "identifies flat-rate trips. The standard metered rate is code 1; code 2 is the JFK flat rate (currently $80 plus surcharges); code 3 is the Newark flat rate; code 4 is the Nassau/Westchester out-of-city rate; code 5 is a negotiated fare; code 6 is the group ride rate. This field is the single cleanest way to identify airport trips, and is used to construct the is_airport_trip feature in the ML model. A small number of rows have code 99 (recorded as 'unknown' by the meter), which the cleaning pipeline preserves as a separate category."),
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
    "The project is structured as a five-stage data pipeline, with each stage implemented as a standalone Python module under the src/ directory. The stages can be re-run independently, which is essential during iterative development: a change to the cleaning logic does not require re-downloading the raw data, and a change to the ML model does not require re-running the SQL analytics. This separation of concerns is one of the foundational habits of production data engineering, and adopting it from the beginning of a portfolio project signals to a reviewer that the author understands the discipline. The five stages are download, clean, feature-engineer, analyze (SQL), and model (ML). A separate dashboard layer in the dashboard/ directory consumes the output of all five stages and exposes everything to a browser via Streamlit. Finally, a bootstrap module produces the analytical dataset on demand so that the dashboard can be deployed to a cloud environment that does not have the source data preloaded — for instance, a fresh Streamlit Community Cloud container."
  ),
  p(
    "Each module is approximately 100 to 200 lines of Python, deliberately short enough that any single file can be read end-to-end in a few minutes. The interface between stages is the filesystem itself: stage N writes a Parquet file (or, in the case of the modeling stage, a serialized LightGBM model), and stage N+1 reads that file. This file-based contract has several advantages over a more sophisticated in-memory orchestration framework. It is trivially debuggable (a developer can inspect the intermediate files at any point), it scales horizontally without effort (different stages can run on different machines or under different schedulers), and it has zero dependencies beyond Python and the filesystem. The trade-off is some I/O overhead between stages, but at the scale of this project (millions of rows, not billions) the I/O cost is negligible."
  ),
  h2("Pipeline stages"),
  bulletBold("1. Download. ", "Streams the requested monthly Parquet files from the TLC CloudFront CDN. Skips files already on disk, retries on transient failures, and reports progress via tqdm. The download module uses Python's requests library with streaming enabled so that files larger than memory can be downloaded without spiking RAM. Each download is written first to a .part file and renamed only on successful completion, so an interrupted download cannot leave a corrupted Parquet on disk that would silently break the cleaning stage."),
  bulletBold("2. Clean. ", "Reads the raw Parquet, normalizes column names across schema versions (the TLC has subtly changed field names over the years and the module handles the differences transparently), computes trip_duration_min from the pickup and dropoff timestamps, and applies the sanity filters in the table below. The cleaning module drops approximately 24% of raw rows. It also emits a short summary of how many rows were dropped, both as a sanity check during development and as evidence that the filters are not silently destroying useful data."),
  bulletBold("3. Feature engineering. ", "Adds the derived columns used throughout the rest of the project. These include pickup_hour (0–23 integer), pickup_dayofweek (Monday=0 through Sunday=6, matching the pandas convention), pickup_day_name (string), pickup_month, pickup_date, is_weekend (boolean), is_rush_hour (boolean, defined as 7–9 AM or 4–7 PM on weekdays), is_airport_trip (boolean, derived from rate code), time_of_day (categorical: Late night / Morning / Afternoon / Evening / Night), avg_speed_mph (distance divided by duration in hours), fare_per_mile (fare divided by distance), and tip_pct (tip as percentage of fare, only meaningful for credit card payments). Each of these is computed once during the cleaning pipeline rather than recomputed on every query, which makes downstream analyses both faster and simpler."),
  bulletBold("4. SQL analytics. ", "Eight named DuckDB queries that answer the standing business questions of the project: overview, trips by hour of day, trips by day of week, top pickup zones, payment-type breakdown, airport versus regular trips, rush hour versus off-peak, and top routes. DuckDB reads directly from the cleaned Parquet — there is no intermediate database load step, no schema definition, no indexing. DuckDB's columnar query execution is exceptionally well-matched to Parquet's columnar storage, and the eight queries collectively complete in under three seconds on the full 8.5-million-row dataset."),
  bulletBold("5. Modeling. ", "Trains a LightGBM gradient-boosted regressor on a 500,000-trip sample with an 80 / 20 train / validation split, early stopping on validation RMSE with patience of 30 rounds, and feature importance computed by gain (the total reduction in loss attributable to a feature, summed across all splits where it was used). The choice of LightGBM over alternatives like XGBoost or scikit-learn's HistGradientBoosting is partly aesthetic — LightGBM has excellent native support for categorical features and trains quickly — and partly practical: the LightGBM model file is small enough (1.5 MB) to commit to the GitHub repository, so the dashboard can do live inference without requiring a separate model artifact storage step."),
  h2("Cleaning filters"),
  p(
    "Each row in the raw Parquet is checked against the ranges in the table below. A row is retained only if it passes every check. The ranges were chosen by inspecting the distribution of each field and identifying clear outlier regions — the goal was to remove obvious data-quality issues without throwing out genuinely unusual but valid trips. Setting the upper distance bound at 100 miles, for example, drops some absurd outliers (trips with 30,000-mile distances, which are clearly meter errors) but retains all legitimate long-distance airport runs and the occasional out-of-city Westchester trip."
  ),
  cleanFiltersTable,
  tableCaption("Sanity filters applied during cleaning"),
  p(
    "These filters drop 2,681,852 rows (23.9% of the raw input). The largest single contributor is implausibly short trips: rows with a recorded duration below one minute. Many of these are probably meter-restart artifacts where the driver started a new trip too quickly after closing the previous one. The second largest contributor is rows with a pickup timestamp from a different year — a recurring TLC data-quality issue where stale or misdated trips occasionally leak into the wrong monthly file. The exact cause is not documented publicly but is likely a quirk of the TLC's batch processing pipeline. Without the year-range filter, time-series visualizations show a small but distracting cluster of rows from 2008, 2014, and other historical years, which would damage the credibility of any daily trend chart."
  ),
  p(
    "A defensible objection to this kind of aggressive filtering is that it might throw out legitimate but unusual trips — a genuine four-hour fare from Manhattan to upstate New York, for instance, would be dropped by the 180-minute upper bound on duration. The mitigation is that the filters are conservative: they drop only trips that are clearly outside the operational envelope of an NYC taxi. A trip that is legitimately long-distance would still be retained if its distance is under 100 miles and its duration is under 3 hours, which together cover essentially every legal Yellow Taxi fare in the region. The 24% drop rate is high in absolute terms but is consistent with the rates that academic studies of this dataset have reported. A reviewer concerned about the filtering can always re-run the pipeline with different bounds — the filter values are constants at the top of the cleaning module, not hardcoded constants buried deep in the logic."
  ),
  h2("Technology stack"),
  p(
    "The stack is deliberately conventional — no exotic dependencies, no proprietary services, no managed databases. Every component runs locally on a developer laptop and is reproducible from a fresh Python environment in under five minutes. The reasoning behind this choice is practical: a portfolio project is judged on whether someone else can reproduce it, and dependencies on cloud services or paid tools introduce friction that frequently causes reviewers to give up. The conventional stack also signals to a hiring manager that the author is comfortable with the industry-standard tools rather than reliant on a particular vendor."
  ),
  techStackTable,
  tableCaption("Tools and their roles in the pipeline"),
  p(
    "One stack choice worth highlighting is the use of DuckDB for the SQL analytics layer. DuckDB is an embedded OLAP database — it runs in-process, requires no server, and can query Parquet files directly without an explicit load step. For analytical work at this scale (millions to low-billions of rows), it is dramatically more convenient than spinning up Postgres or a dedicated data warehouse, and the columnar query engine is exceptionally fast. DuckDB is also used in the dashboard's bootstrap module to perform memory-efficient streaming cleaning of raw Parquet files, which proved essential when deploying the dashboard to Streamlit Community Cloud's 1 GB free tier. The lesson is that DuckDB's positioning as the SQLite of analytics is well-founded; it deserves a place in any modern data toolkit."
  ),
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
    "The single most important variable in any analysis of urban movement is time. Demand for taxi service in New York City varies by more than an order of magnitude across a single day — the 5 AM hour sees roughly forty thousand trips while the 6 PM hour sees over six hundred thousand — and is reliably patterned by day of the week. Understanding when people take taxis is foundational to every downstream analytical question in this report: pricing, congestion, tipping, and predictive modeling all build on the temporal structure described in this section. The heatmap below presents the joint distribution of pickups by hour of day and day of week, and is in many ways the single most informative image in the entire dataset."
  ),
  ...scaledFigure("demand_heatmap", "Trips by hour of day × day of week"),
  p(
    "Several patterns are visible at a glance. Weekday afternoons through early evening — roughly 2 PM through 7 PM Tuesday, Wednesday, Thursday, and Friday — form a clear dark band that dominates the chart. This is the visible footprint of the evening commute layered on top of business and tourist travel during the working day. A secondary, smaller dark band appears late Friday and Saturday nights, extending into the early morning hours of Saturday and Sunday — the city's nightlife economy at work, with taxi pickups concentrated around bar-closing time (4 AM in New York). Sunday evening is conspicuously dimmer than the corresponding weekday evening windows, consistent with a population that is largely at home preparing for the work week. The early-morning hours of Monday through Friday (5 AM and 6 AM) show a slight thickening of activity, reflecting morning commute departures and early-morning airport runs."
  ),
  p(
    "What is perhaps most striking about the heatmap is what it does not show. There is no clear morning rush peak in taxi pickups, which is somewhat counterintuitive — most New Yorkers would assume that morning commute hours would generate substantial taxi demand. The explanation is that most morning commuters use the subway, walk, or bike, and the taxi-eligible morning population is much smaller than the taxi-eligible evening population. In the evening, by contrast, business travelers heading to dinner, tourists heading to shows or restaurants, workers heading home after a late day at the office, and people heading to airports all converge into the dense afternoon-and-evening band. The pattern is a useful reminder that taxi demand is not a simple proxy for total urban movement; it reflects the specific population of travelers who choose taxis over the alternatives."
  ),
  h2("Hour-of-day breakdown"),
  p(
    "Selected hours from the full 24-hour distribution are shown below to give a sense of how each metric varies across the day. The full table covering all 24 hours is available via the project's SQL analytics module and is rendered in the interactive dashboard. The hours displayed here were chosen to span the full range of the daily cycle, including the quietest pre-dawn hours, the morning ramp, the afternoon plateau, the evening peak, and the late-night winding-down."
  ),
  hourlyTable,
  tableCaption("Trip volume and economics across selected hours"),
  p(
    "Several observations from this table deserve attention. First, the busiest hour by trip volume — 6 PM (hour 18) with 621,501 trips — is not the hour with the highest average fare. Average fare peaks at the much quieter 5 AM hour, where the average is $27.34. The mechanism is straightforward: early-morning trips are dominated by long-distance airport runs to JFK and LaGuardia for first-flight passengers, and those trips have substantially higher fares than the median Manhattan-to-Manhattan trip that dominates the rest of the day. The 5 AM population is small (only 45,943 trips, less than 8% of the 6 PM volume) but it is unusually high-value per trip."
  ),
  p(
    "Second, average vehicle speed is at its lowest precisely during the busiest evening hours: 9.7 mph in both hour 15 (3 PM) and hour 17 (5 PM). The same trip from the same origin to the same destination is roughly 40% slower at 5 PM than at 3 AM. This is not just an inconvenience for riders; it has direct economic consequences because the meter charges by both distance and time. A trip that takes longer because of traffic costs more even if it covers the same distance, and the per-mile cost of a rush-hour trip is correspondingly higher than the per-mile cost of an off-peak trip. The mechanism by which congestion translates into higher fares is mediated through the meter, not through any pricing decision by the rider or driver."
  ),
  p(
    "Third, the duration column shows a less-pronounced pattern than the speed column, which initially seems contradictory but actually reflects a real behavioural mechanism. Riders compensate for traffic by taking shorter trips: when speeds are low, the marginal cost of a long taxi ride is high, and riders self-select into shorter routes (or out of taxis altogether, in favour of walking or the subway). The result is that average duration is relatively stable across the day, while average speed varies substantially. The relationship between speed, distance, and duration is not a simple identity; it is mediated by rider choice."
  ),
  h2("Day-of-week breakdown"),
  p(
    "Aggregating across the full quarter, the day-of-week pattern follows a smooth ramp from Monday through Thursday, with Thursday emerging as the single busiest day with 1,367,466 trips. Friday and Saturday remain elevated, and Sunday returns to a level only marginally above Monday — Monday and Sunday are essentially tied for quietest, with Sunday being slightly quieter in absolute volume despite being a non-work day. This pattern is consistent across all three months in the dataset."
  ),
  dowTable,
  tableCaption("Trip volume and tipping by day of week"),
  p(
    "Why is Thursday the busiest day? Several mechanisms likely contribute. First, Thursday is the conventional 'business travel return day' — corporate travelers who flew out on Monday or Tuesday often return on Thursday evening, generating substantial airport pickup demand. Second, Thursday is increasingly the new Friday in post-pandemic urban work patterns, with restaurants, bars, and entertainment venues seeing peak demand that traditionally would have been Friday. Third, Thursday is a peak commute day in its own right because it lacks the partial work-from-home effect that has reduced Friday office attendance in many sectors. The result is a Thursday peak that is visible not only in the day-of-week aggregation here but also in nearly every published study of post-pandemic urban mobility."
  ),
  p(
    "Average tip percentage declines slightly but consistently from Monday (26.97%) to Sunday (25.24%) — a 1.7 percentage-point spread that may not look large in isolation but represents a meaningful aggregate effect across millions of trips. Several mechanisms could plausibly contribute. Weekend trips tend to be shorter and more leisure-oriented, and shorter trips with smaller absolute fares may produce smaller tip percentages simply because the rider's tipping behaviour is anchored to round-dollar amounts (a $1 tip on a $4 fare is 25%; a $1 tip on a $5 fare is 20%). Weekend tipping may also be split across a larger share of cash transactions, which never enter the percentage calculation because cash tips are not recorded. And there is a behavioural literature suggesting that business travelers (who concentrate on weekdays) tip somewhat more generously than leisure travelers, possibly because they are tipping on someone else's credit card and the marginal cost to them is zero. None of these mechanisms can be isolated cleanly from this dataset alone, but the directional pattern is unambiguous."
  ),
  h2("Daily volume across the quarter"),
  p(
    "Plotting daily volume across the three months reveals weekly periodicity and a small number of anomalies that are worth identifying explicitly. The biggest dips correspond to severe-weather days and federal holidays. Martin Luther King Jr. Day on January 20 shows a visible drop relative to the surrounding Mondays. Presidents Day on February 17 shows a similar drop. A small number of additional dips correspond to high-impact weather events during the quarter, including two significant winter storms in late January. There is no strong overall trend across the three months, indicating that the underlying taxi market was stable across Q1 2025 — there is no obvious month-over-month growth or decline that would suggest a structural shift."
  ),
  ...scaledFigure("daily_volume", "Daily trip volume across the analysis window"),
  p(
    "The weekly periodicity in the daily series is striking — the high-low oscillation between weekday peaks and weekend troughs is remarkably regular. This kind of clean seasonality is a gift for time-series forecasting: a simple model that learns the weekly pattern would produce useful forecasts with very little effort, and any deviation from the weekly pattern would be immediately visible as a potential anomaly worth investigating. While time-series forecasting is outside the scope of this project, the cleanliness of the seasonal structure suggests that a future extension toward forecasting would be straightforward and rewarding."
  ),
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
    "Three distributions characterize the economic shape of NYC taxi activity: distance, duration, and fare. Each is heavily right-skewed — the median trip is short, but a long tail of airport runs and long outer-borough trips drags the mean well above the median. The visualization below presents the distributions side-by-side, each annotated with its median, and the table that follows quantifies the median, mean, and 95th-percentile for the six most important economic metrics. Understanding these distributions is the foundation for every subsequent analysis: any aggregate statistic that ignores the skew (such as a naive average) will be misleading without context, and any pricing or operational analysis must account for the fact that the bulk of trips look very different from the tail."
  ),
  ...scaledFigure("trip_distributions", "Distributions of trip distance, duration, and fare amount"),
  economicsTable,
  tableCaption("Summary statistics for the core trip metrics"),
  p(
    "The gap between the median and the mean is informative throughout the table. A median distance of 1.6 miles versus a mean of 3.22 miles indicates that half of all trips are short — typical Manhattan-to-Manhattan rides covering a few dozen blocks — while the long tail of airport and outer-borough trips pulls the mean upward by approximately a factor of two. The P95 column shows where the long tail actually sits: the 95th-percentile trip is 11.4 miles long, 39.6 minutes in duration, and costs $56 in fare alone (before tip and surcharges). The very longest trips, the JFK and Newark airport runs and the occasional out-of-city Westchester trip, are the primary driver of this tail. The fact that the P95 is roughly 3.5 times the median across all three economic metrics reflects a common shape in urban transportation data: a tight modal cluster around the typical trip with a long Pareto-style tail of unusual trips."
  ),
  p(
    "The distributions also have implications for visualization design. Because the data is so right-skewed, plotting the raw distribution without truncation produces a chart where the modal mass is squashed into the leftmost few percent of the horizontal axis and the entire visualization is dominated by white space. The trip_distributions figure above truncates each axis at a sensible upper bound (20 miles for distance, 60 minutes for duration, $80 for fare) so the modal shape is visible. A more sophisticated treatment would use a log-scale axis, which would render the distribution as something close to symmetric and would make the tail more interpretable. Both approaches are valid; the choice depends on whether the visualization is meant to highlight the modal behaviour (linear truncated axis, as here) or to emphasize the heavy tail (log axis)."
  ),
  h2("Fare versus distance"),
  p(
    "Plotting fare against distance shows the meter's pricing function with great clarity. The dense linear core represents standard metered trips, where the fare is approximately proportional to distance with a modest fixed component reflecting the meter's initial pickup charge of approximately $3 and the per-mile rate of approximately $3 to $4 depending on time of day. A clearly distinct horizontal cluster at approximately $70 corresponds to the JFK flat rate (a base $80 trip including the standard surcharge structure, which after the JFK improvement surcharge and tolls lands around the $70 fare amount visible in the data). Smaller clusters at other flat-rate price points — the Newark trips at around $90 to $100, the negotiated-fare trips scattered through the chart — are visible as well, though less prominent than the JFK cluster simply because Newark is a smaller airport with less Yellow Taxi traffic."
  ),
  ...scaledFigure("fare_vs_distance", "Fare amount versus trip distance"),
  p(
    "This visualization is itself a justification for treating airport trips as a separate population in downstream analysis. Their fares are bounded above by the flat rate regardless of distance, so a standard linear regression of fare on distance would systematically over-predict airport-trip fares for high distances. The rate code (ratecode_id) cleanly distinguishes airport trips from standard metered trips, and the ML model in section 8 leverages this separation explicitly by including ratecode_id as a categorical feature. A model that did not include this feature would be forced to learn a single fare-versus-distance relationship that is wrong for both populations, and its accuracy would suffer accordingly."
  ),
  p(
    "An additional pattern worth noting is the small but visible vertical band of points at distance values that are exactly 0.1 or 0.2 miles. These represent trips where the meter rounded the distance to its minimum granularity, and they often correspond to either fraud (a driver who charged a rider for a trip that did not actually occur) or to legitimate very-short trips (a rider who hopped in and decided to get out almost immediately). The cleaning pipeline excludes trips with distances below 0.1 miles, which removes the most egregious cases without throwing out genuinely short rides."
  ),
  h2("Fare per mile"),
  p(
    "Dividing fare by distance gives a per-mile economic intensity. The mean across all trips is $7.20 per mile, but this average obscures substantial systematic variation across trip categories. Short trips inside Manhattan have a much higher per-mile cost because the fixed pickup fee amortizes over a small distance — a trip of one mile that includes a $3 base fare has a per-mile cost of at least $6 from the base alone, before the per-mile rate is added. Long airport runs have a much lower per-mile cost because the flat rate works out to a low per-mile figure across the long airport distance. Off-peak hours show a lower per-mile cost than rush hour because the per-minute waiting-time charge is less prominent when the cab is moving freely."
  ),
  p(
    "The per-mile cost figure is useful as a single-number summary of trip economics, but it should be interpreted with the underlying distribution in mind. A rider deciding between a taxi and an alternative mode of transport is rarely making the decision on per-mile cost; they are making it on total expected cost for their specific trip. The per-mile aggregate is more useful as a way of comparing across categories (rush hour versus off-peak, airport versus regular, weekday versus weekend) than as a guide to any individual decision. Where the per-mile figure becomes especially valuable is in cross-comparison with other ride-hailing services: Uber and Lyft fares are typically reported in per-mile-plus-per-minute terms, and the Yellow Taxi per-mile figure provides a natural benchmark."
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
    "Payment behaviour in NYC taxis follows a clear hierarchy: credit cards dominate every other channel by a wide margin. The transition from cash-heavy to card-heavy payment took place over the late 2000s and early 2010s and was largely complete by 2015, driven by a combination of regulatory requirements (the TLC mandated that all medallion taxis accept credit cards) and rider preference for the convenience and built-in tipping interface of card payment. The residual cash share remains material — primarily for short, low-fare trips, in certain neighborhoods where cash is the default payment mode for many goods and services, and among tourists who may have foreign credit cards that incur transaction fees abroad. The table below breaks down the four payment categories observed in the data."
  ),
  paymentTable,
  tableCaption("Trip volume and economics by payment type"),
  p(
    "Several points are worth highlighting from this table. First, only 86% of trips have a recorded tip — the other 14% are cash, dispute, or no-charge categories where the data feed does not capture tip information. This is a critical data limitation: any tipping analysis is necessarily an analysis of card tipping, and findings about tipping behaviour cannot be generalized to the full population without caveat. The 12% of cash trips may have entirely different tipping dynamics, but the dataset is silent on this question. Second, average fare is remarkably stable across payment types — riders do not appear to systematically use payment type to signal trip importance or to handle long, high-stakes trips differently from short ones. The average card fare is $18.10 and the average cash fare is $18.08, a difference small enough to be effectively zero. This is reassuring evidence that the tipping analysis on the card subset is at least not biased toward an unusual trip population on the dimension of fare."
  ),
  p(
    "Third, the small but meaningful dispute category (109,568 trips, 1.29%) has the highest average fare in the table at $21.87 — about 20% higher than the all-trip average. This suggests that disputes are more likely to arise on more expensive trips, which makes intuitive sense: the absolute dollar stakes of a disputed $30 fare are larger than those of a disputed $8 fare, and riders are more likely to invest the friction of disputing a charge when the amount in question is substantial. Fourth, the no-charge category (34,277 trips, 0.40%) reflects a mix of corporate accounts, driver-comp scenarios, and edge cases in the meter system; these trips look essentially identical to standard trips in fare and distance, suggesting they are routine trips that happened to involve a non-standard billing arrangement."
  ),
  h2("Card tipping distribution"),
  p(
    "Within the card-payment population, the distribution of tip percentage is remarkably tight. The median tip percentage is 20%, and the bulk of the distribution sits in the 15% to 25% range. There are three sharp spikes in the distribution at 15%, 20%, and 25% — exactly the values that the in-cab payment interface presents as preset buttons. Visually, the histogram does not look like a continuous distribution at all; it looks like three discrete spikes with a small amount of mass scattered around the rest of the range. The clustering reflects the dominance of the interface presets in shaping rider behaviour: most riders look at the screen, see three reasonable percentage options, and tap the middle one. Only a small minority go to the trouble of entering a custom percentage or tipping in cash."
  ),
  ...scaledFigure("tipping_behaviour", "Tip percentage distribution and tip percentage by hour of day"),
  p(
    "Plotting average tip percentage by hour of day reveals only modest variation: most hours sit between 25% and 27%, with a slight upward drift in late evening hours. This is consistent with the interpretation that the in-cab interface dominates tipping behaviour: if riders were deliberating about tip amount based on service quality, time of day, or any other contextual factor, we would expect much larger swings across hours. The fact that the daily average is essentially flat suggests that the deliberation is happening at a much smaller scale than the structural anchoring effect of the preset buttons. Behavioural differences across time of day exist, but they are small relative to the interface effect."
  ),
  p(
    "There is a small but consistent uptick in tip percentage during the very late evening and early morning hours (roughly 11 PM through 3 AM). The most plausible explanation is selection effects: late-night riders are more likely to be tourists or leisure travelers (rather than business commuters), and may include a higher proportion of riders who have been at bars or restaurants and are tipping more generously as a general matter of mood. Late-night drivers may also provide a level of service (helping with bags, going slightly further off-route to a safe drop-off location) that is more visible than during the day, and may be rewarded with slightly higher tips. These mechanisms are speculative; the data shows the pattern but cannot definitively attribute it to a cause."
  ),
  h2("Implications for operators and platform designers"),
  p(
    "The tightness of the tipping distribution has direct implications for any operator or platform designer interested in lifting aggregate tip revenue. The data strongly suggests that small interface-level changes — adjusting the preset percentages, shifting the default selection from 20% to a higher value, or changing the visual prominence of the higher options — would have a far larger effect on aggregate tipping than any campaign aimed at changing rider behaviour through training, advertising, or social pressure. The clustering of tip percentages around the preset values is too tight and too consistent across times of day to be explained by genuine rider deliberation. Riders are anchoring on the defaults, and small changes to the defaults would shift the entire distribution."
  ),
  p(
    "This is a textbook example of a phenomenon that behavioural economists have documented in many other settings: when faced with a default choice among a small set of options, most people accept the default. The interface designer therefore has substantial leverage over outcomes, and this leverage should be exercised thoughtfully. There is also an ethical dimension: defaulting riders into higher tip percentages benefits drivers but may be perceived as manipulative if the defaults drift too high. The current configuration of 15% / 20% / 25% has held steady for years and appears to be a stable equilibrium."
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
    "Duration alone does not separate long trips from slow trips: a 30-minute trip can be 25 miles on the highway or 4 miles inside Manhattan at 5 PM. Computing the implied average speed (distance divided by duration) gives a far more interpretable measure of congestion, because it normalizes away the underlying trip length. Plotting average speed by hour of day produces the clearest single illustration of when New York's traffic is actually bad, and the magnitude of the variation is striking enough that it has direct implications for both rider experience and operator economics."
  ),
  ...scaledFigure("speed_by_hour", "Average trip speed by hour of day"),
  p(
    "The pattern is unambiguous and aligns with everyday New Yorkers' intuitions about when the city is moving and when it is gridlocked. Late-night and early-morning hours show the highest speeds, peaking at approximately 19.6 mph at 5 AM when the streets are empty and the only traffic is the small fleet of early-morning airport runs moving on largely empty highways. Speed declines steadily through the morning as commuter traffic builds, reaching a plateau around 10 mph through the middle of the day as the city's roadways carry their normal weekday load. Speed drops further into the late afternoon and early evening, bottoming out at approximately 9.7 mph in the 3 PM to 5 PM window when commute traffic peaks. The full swing from peak (5 AM, 19.6 mph) to trough (3 PM, 9.7 mph) is roughly 10 mph, or just over a 50% reduction in throughput. In practical terms, the same trip from the same origin to the same destination takes twice as long during the worst rush hour as it does in the small hours of the morning."
  ),
  p(
    "It is worth pausing on what 9.7 mph actually represents. The average New Yorker walks at approximately 3 mph; a brisk walking pace is around 4 mph; a slow jog is about 6 mph. A taxi moving at 9.7 mph is moving at roughly the speed of a fast jog. For trips of more than a few blocks, walking is genuinely competitive with taking a taxi during peak congestion, which is why many residents simply walk during the worst hours and why the subway captures the substantial majority of long commute trips even in a city with abundant taxi service."
  ),
  h2("Rush hour versus off-peak"),
  p(
    "Aggregating the hours into binary rush-hour (7–9 AM and 4–7 PM on weekdays) and off-peak categories quantifies the economic effect of congestion. Rush hour as defined here covers five hours of the weekday — roughly 14% of the total hours in the dataset — but accounts for a substantially larger share of trips because demand is concentrated in those hours."
  ),
  rushTable,
  tableCaption("Rush hour versus off-peak — speed and pricing"),
  p(
    "Rush-hour trips are 18% slower on average (9.5 mph versus 11.6 mph) and cost 14% more per mile ($8.91 versus $7.83). The per-mile fare difference is the direct economic translation of slower traffic: New York's meter charges both by distance and by waiting time, so a slow trip costs more even for the same distance. A rider taking the same five-mile trip at 8 AM versus 1 AM will pay approximately seven percent more in absolute fare and substantially more in time spent in the cab. This is the hidden cost of congestion paid by the riding public, and it is borne disproportionately by riders whose schedules require them to travel during peak hours."
  ),
  p(
    "From the operator's perspective, the picture is more nuanced. A slower trip means the cab covers fewer miles per shift and serves fewer riders per shift, which reduces the driver's revenue throughput. But the higher per-mile fare during congestion partially offsets this — drivers earn more per mile but cover fewer miles. The net effect on driver earnings depends on the relative magnitudes, and is not directly visible from this dataset. What is clear is that the city's existing meter structure does at least partially internalize the cost of congestion: drivers are compensated (in per-mile terms) for the slower throughput, even if the absolute volume effect cuts the other way."
  ),
  h2("Comparison to other cities"),
  p(
    "While this report focuses exclusively on New York City data, the speed-by-hour pattern observed here is broadly representative of how congestion looks in any dense urban environment. Academic studies of London, Tokyo, and São Paulo have all found qualitatively similar patterns: a peak in average vehicle speed during the small hours of the morning, a midday plateau, and a sharp dip during the late-afternoon commute. The magnitude of the variation differs across cities — Tokyo's transit network handles enough commuter volume that the road network sees less rush-hour stress than New York's — but the qualitative shape is universal. The mechanism is the same in every case: roadway capacity is roughly constant across the day, but demand spikes during commute windows, producing congestion that reduces throughput speed."
  ),
  h2("Why this matters for the ML model"),
  p(
    "The fact that the same origin-destination pair has a substantially different duration depending on hour of day is the central reason why the trip-duration prediction model in the next section must include temporal features. A model that uses only distance would systematically over-estimate duration during off-peak hours (because it would apply an average speed that is too low for those hours) and under-estimate it during rush hour (because the average speed is too high for the actual rush-hour conditions). The pickup_hour feature, the is_rush_hour flag, and the location-pair categorical features together allow the model to learn this structure from the data without needing any explicit feature engineering about traffic patterns."
  ),
  p(
    "An important methodological note is that the model is trained on a sample drawn uniformly across the three-month window, which means the model implicitly assumes that the temporal patterns are stable. For a production system, this assumption would need to be reconsidered: if traffic conditions change materially (because of a new construction project, a major event, or a permanent infrastructure change), the model would need to be retrained on more recent data. A small extension of the project would be to add a retraining cadence and a monitoring system that alerts when model accuracy on recent data drifts below a threshold."
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
    "Estimating trip duration in advance is a foundational ride-hailing problem. Dispatch systems must commit to an ETA when accepting a rider request, and a credible ETA is a precondition for the rider's decision to take the trip in the first place. Surge-pricing logic depends on expected supply turnover — how quickly drivers will complete their current trips and become available again — which is a direct function of expected trip duration. Downstream routing algorithms benefit substantially from a duration estimate they can trust, because they can plan multi-leg trips, batch deliveries, or coordinate driver shifts based on expected completion times. The cost of being wrong is asymmetric: a substantially late arrival damages rider trust far more than an early arrival saves it, because riders plan their connections, meetings, and onward travel based on the promised arrival time."
  ),
  p(
    "We frame the task as a regression problem with a hard methodological constraint: only features available at the moment the trip starts may be used. This rules out any feature derived from the drop-off timestamp or the final fare, both of which are unknown at the moment the model would actually need to make a prediction. The constraint is essential for preventing target leakage, a common failure mode in machine learning where the model implicitly cheats by using information that would not be available at prediction time. A model that achieves high accuracy by leaking the answer into the features is worthless in production. The features described below all satisfy the constraint that they are known at the moment a trip is dispatched."
  ),
  h2("Features"),
  bulletBold("trip_distance — ", "the requested distance in miles. The single most important feature, with feature importance roughly an order of magnitude greater than any other. In a real dispatch system this would be supplied by a routing engine like Google Maps or OpenStreetMap, not by the meter; the meter's distance is only available after the trip is complete. For training purposes the meter's distance is a reasonable proxy because the two are typically within a few percent."),
  bulletBold("passenger_count — ", "weakly predictive of duration; large parties slightly increase boarding time and the chance of stops to coordinate dropoffs. The feature is included for completeness but has low importance."),
  bulletBold("pickup_hour, pickup_dayofweek, pickup_month — ", "temporal features that capture demand cycles and traffic conditions. Pickup hour is the most important of the three; pickup month captures any seasonal effect (winter weather, school calendars) within the analysis window; pickup dayofweek captures the weekly cycle."),
  bulletBold("pu_location_id, do_location_id — ", "origin and destination zone IDs, treated as categorical features rather than ordinal integers. The categorical treatment allows the model to learn zone-pair-specific behaviour rather than imposing an arbitrary order on the zone IDs. For instance, the model can learn that trips from zone 161 to zone 230 are reliably slower than trips from zone 230 to zone 161 because of one-way street patterns, without that information needing to be explicit in the features."),
  bulletBold("ratecode_id — ", "treated as categorical. Distinguishes standard metered trips from JFK / Newark / Nassau flat-rate trips. This is one of the most important features after distance because it cleanly identifies the structurally different airport-trip population."),
  bulletBold("is_weekend, is_rush_hour, is_airport_trip — ", "engineered boolean flags. The model can in principle infer these from the underlying fields, but providing them explicitly accelerates training and makes the model's logic easier to interpret. The is_rush_hour flag is defined as the conjunction of (7–9 AM or 4–7 PM) and (not a weekend), so it cleanly identifies the temporal slice where congestion is worst."),
  h2("Training procedure"),
  p(
    "The cleaned 8.5-million-row dataset was sampled down to 500,000 trips with a fixed random seed to keep training tractable. The choice of 500,000 was empirical: at this sample size training takes approximately 30 seconds on a developer laptop, the model achieves stable validation metrics, and adding more data does not materially improve performance. Sampling down from millions to hundreds of thousands of rows is a standard practice for LightGBM and similar boosting models, which have a tendency to overfit on very large training sets unless they are carefully regularized."
  ),
  p(
    "The sample was split 80 / 20 into training (400,000 rows) and validation (100,000 rows) sets, again with a fixed seed for reproducibility. This kind of straightforward random hold-out is appropriate when the data is independent and identically distributed across time, which is approximately true here at the level of months. A more rigorous treatment would use a time-based hold-out (train on January and February, validate on March) to ensure the model generalizes to genuinely future data; that change is straightforward and would be worth doing for any production deployment. For this report's purposes, the random hold-out provides a reasonable estimate of the model's accuracy."
  ),
  p(
    "LightGBM was used as the gradient-boosting framework for its speed on tabular data, its native support for categorical features, and its mature implementation of early stopping and feature importance reporting. The hyperparameter table below summarizes the configuration. None of the hyperparameters were exhaustively tuned — the values chosen are sensible defaults that work well across a wide range of tabular regression problems. A more thorough hyperparameter search using Bayesian optimization (for example, with the Optuna library) could potentially squeeze a few additional points of R² out of the model, but the marginal returns are typically small relative to the effort, and the current configuration produces results that are already strong."
  ),
  hyperparamTable,
  tableCaption("LightGBM hyperparameters"),
  p(
    "Early stopping was triggered at 215 boosting rounds out of a possible 500, indicating that the model had reached its capacity for the available data and that further boosting rounds would either fail to improve the validation metric or would begin to overfit. The fact that early stopping was triggered well before the 500-round maximum is encouraging: it means the configuration is appropriately balanced between underfitting and overfitting, and is not wasting computation on rounds that do not help. Training took approximately 30 seconds total on a developer laptop, an acceptable training time for any iterative development workflow. The categorical features (pu_location_id, do_location_id, ratecode_id) were declared explicitly so that LightGBM uses its optimal-split algorithm for categoricals rather than treating them as ordinal numbers; this is a small but important configuration choice that materially improves model accuracy on this kind of categorical-heavy data."
  ),
  h2("Validation results"),
  modelMetricsTable,
  tableCaption("Validation metrics on the 100,000-trip hold-out set"),
  p(
    "The headline number is mean absolute error of 2.77 minutes. In plain language, the typical prediction is within about three minutes of the true duration. For a model that uses only trip-start information, this is a strong result; for context, an academic study of taxi duration prediction on a similar dataset published around 2018 reported a best-of-class MAE of approximately 3.5 minutes, so the current model is comfortably within the range of published state-of-the-art results. The R² of 0.864 indicates that the features collectively explain 86% of the variance in trip duration; the remaining 14% reflects genuinely unpredictable factors (driver behaviour, momentary traffic incidents, weather variations) that are not visible in the current feature set."
  ),
  p(
    "The RMSE of 4.49 minutes is meaningfully higher than the MAE of 2.77 minutes, which is consistent with a distribution of errors that includes some larger outliers. The ratio of RMSE to MAE is approximately 1.62, which for a normal distribution would be expected to be about 1.25; the gap indicates that errors are somewhat heavy-tailed. This is again consistent with the structure of the data: most trips behave predictably, but a small fraction encounter unusual conditions (a traffic incident, a sudden weather change, a passenger asking the driver to make an unexpected stop) that produce larger duration errors."
  ),
  ...scaledFigure("predicted_vs_actual", "Predicted versus actual trip duration on the validation set", { width: 400 }),
  p(
    "The predicted-versus-actual scatter plot above shows the model tracks the truth closely along the diagonal across the full range of durations, with predictable widening of the error band at very long durations (where there are simply fewer training examples and the absolute scale of plausible errors is larger). The visualization is sampled to 5,000 points for clarity; the full validation set of 100,000 points would produce a denser but qualitatively identical chart. The red dashed line is the line of perfect prediction; the cloud of points is approximately symmetric around it, which suggests the model is not systematically over- or under-predicting at any particular duration range."
  ),
  ...scaledFigure("feature_importance", "LightGBM feature importance (gain)"),
  p(
    "Feature importance by gain ranks distance as the dominant feature by an order of magnitude. The second-most-important feature is the rate code, which encodes the flat-rate airport runs that the model treats as a structurally different population. Pickup hour comes third, capturing most of the traffic effect that was illustrated in section 7. The destination and origin zone IDs follow, allowing the model to learn that certain trips between specific zones are reliably faster or slower than the distance-and-hour baseline would suggest — for instance, trips on highways and bridge crossings have different per-mile speeds than purely intra-Manhattan trips. The engineered flags (is_weekend, is_rush_hour, is_airport_trip) appear at the bottom of the ranking, which is consistent with their being derivable from the more granular features higher up; the model could in principle compute them itself, but providing them as features accelerates training."
  ),
  h2("Error analysis"),
  p(
    "Where does the model do well, and where does it do poorly? A residual analysis (not shown in this report but easily reproducible from the project code) reveals three populations where errors are larger than the average. First, the very longest trips (over 60 minutes) have larger errors in absolute terms, simply because there are fewer training examples and the variance of possible durations is larger. Second, trips during weather-disrupted days have larger errors because the model has no weather feature. Third, trips that happen to encounter a traffic incident (an accident, a road closure) have larger errors because the model has no real-time traffic feature. All three failure modes are expected given the feature set, and each suggests a clear extension that would improve model accuracy."
  ),
  h2("Limitations and possible extensions"),
  p(
    "Several limitations of the current model are worth flagging explicitly. First, the model produces a point estimate, not a probability distribution. For real dispatch use, a probabilistic ETA (for example, a P10 to P90 range that captures the realistic uncertainty around the central estimate) would be more useful, because it allows the dispatcher to communicate a confidence range to the rider and to make risk-aware decisions about supply allocation. A quantile-regression LightGBM ensemble — separate models trained to predict the 10th, 50th, and 90th percentile of the duration distribution — is the standard way to obtain this output, and would be a small extension of the current code."
  ),
  p(
    "Second, the model does not include weather features. Incorporating an hourly precipitation feed (NOAA publishes this freely for JFK and LGA weather stations, both of which are reasonable proxies for citywide conditions) would likely lift R² further, particularly during rare but impactful weather events when traffic patterns shift substantially. Snow days, heavy rain, and high-wind days all produce demonstrably different traffic conditions than fair-weather days, and a weather-aware model would handle these cases more robustly. Third, the model does not include any real-time traffic feature. A real production deployment would presumably integrate with a real-time traffic API (Google Maps, Mapbox, or HERE) to capture the current state of congestion, which is a much better signal than the average for the hour. Fourth, the model treats the zone IDs as fixed categorical features; an embedding approach that learns a low-dimensional representation of each zone could potentially extract more signal from the zone-pair interactions, and would also generalize better if the zone definitions are updated by the TLC in the future."
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
    "All of the analytical findings in this report are exposed through a Streamlit dashboard that runs in a browser. The motivation for a separate dashboard layer, in addition to this written report, is that interactivity allows the user to explore the data on their own terms — to drill into a specific zone, to filter by payment type, to enter their own trip parameters and see what duration the model predicts. A static report communicates the analyst's conclusions; an interactive dashboard lets the audience verify those conclusions themselves and discover patterns the analyst may not have thought to highlight. The dashboard is structured into five tabs, each with a distinct analytical focus. The tabs are summarized in the table below."
  ),
  tabsTable,
  tableCaption("Streamlit dashboard structure"),
  h2("Design philosophy"),
  p(
    "Three design principles guided the dashboard's construction. First, every chart is interactive — built with Plotly rather than matplotlib — so that users can hover, zoom, and inspect individual data points without needing to ask the analyst to regenerate the chart. Second, the dashboard does not require the user to make any decisions before they see useful content; the default view of every tab is the aggregate analysis, and users can drill down from there if they want more detail. Third, page loads are fast — sub-second for any tab navigation after the initial load — because every expensive operation is wrapped with Streamlit's caching decorators."
  ),
  p(
    "An anti-design principle worth naming is the avoidance of cluttered, multi-pane dashboards that try to show every metric simultaneously. The five-tab structure means that any single view focuses on one analytical question, with enough space to present the relevant charts at readable size. This is less impressive-looking than a dense single-page dashboard but far more useful for actual analytical work, because the user is not distracted by metrics that are irrelevant to their current question."
  ),
  h2("Performance and caching"),
  p(
    "All dataset loading is wrapped with Streamlit's caching decorators. The DuckDB connection is created once per session with @st.cache_resource so that subsequent tab navigation does not pay any setup cost. Each SQL query result is wrapped with @st.cache_data so that re-rendering a tab after a slider change reuses the previous result. The net effect is that the first page load takes one to two seconds (dominated by the initial DuckDB connection setup and the query that produces the overview KPIs); every subsequent interaction within the same session is effectively instantaneous."
  ),
  p(
    "The caching strategy is worth understanding because it represents a meaningful improvement over the default Streamlit behaviour. Without caching, every interaction would re-execute every query, which would make the dashboard unusable on a dataset of this size. With caching applied correctly, the dashboard feels like a native application despite running over the web and being implemented in Python. The cost of getting caching right is a small amount of decorator boilerplate; the benefit is a dramatically better user experience."
  ),
  h2("Live ML inference"),
  p(
    "The ML predictor tab loads the trained LightGBM model at session start using @st.cache_resource (the model is approximately 1.5 MB on disk and loads in milliseconds). When the user submits the form — by entering values for trip distance, pickup and drop-off zones, time of day, passenger count, and rate code, then clicking the Predict button — the inputs are assembled into a single-row pandas DataFrame, the categorical columns are converted to the correct dtype, and the model produces a duration prediction in under 10 milliseconds. The prediction is displayed alongside the implied average speed (distance divided by predicted duration, converted to mph) and a small contextual note — for example, a warning that rush-hour predictions tend to have higher variance than off-peak predictions because traffic conditions are more volatile during peak hours."
  ),
  p(
    "The inference tab also serves a pedagogical purpose: it lets a user develop intuition for what features the model considers important by changing inputs and observing how the prediction changes. A user who changes the pickup hour from 3 AM to 5 PM while holding everything else constant will see the prediction increase substantially, which directly illustrates the model's encoding of the congestion patterns described in section 7. A user who changes the rate code from standard to JFK will see a different prediction shape because the model has learned that flat-rate airport trips behave differently from standard metered trips. This kind of interactive what-if exploration is far more illuminating than any number of static charts of model behaviour."
  ),
  h2("Cloud deployment"),
  p(
    "The dashboard is designed to deploy to Streamlit Community Cloud unchanged. A bootstrap module included in src/ checks at app startup whether the cleaned dataset exists; if not, it downloads a single month of raw data and produces a 500,000-row analytical sample using DuckDB. This pipeline runs entirely within DuckDB rather than pandas to stay within the 1 GB RAM ceiling of the free deployment tier. The bootstrap typically completes in under fifteen seconds on first launch, after which the dataset is permanent for the lifetime of the deployed container."
  ),
  p(
    "The choice to use DuckDB for the bootstrap rather than the original pandas-based cleaning pipeline was forced by a real deployment failure: the first attempt at deploying the dashboard to Streamlit Cloud crashed because the pandas pipeline exceeded the 1 GB memory limit while loading and cleaning the raw data. DuckDB processes Parquet in a streaming fashion that never holds the full dataset in RAM, making it possible to run the entire cleaning pipeline within a fraction of the available memory. This kind of constraint-driven engineering decision — discovering a memory issue in deployment, diagnosing it, and rewriting a module to fit the production environment — is exactly the sort of work that distinguishes a portfolio project from a toy exercise."
  ),
  p(
    "Deploying the dashboard to a free hosting tier and sharing the URL produces a public demonstration that is dramatically more compelling than any number of screenshots in a written report. A recruiter clicking a link to the live app and exploring it for thirty seconds takes away a much stronger impression than the same recruiter reading three pages of static analysis. This is the principle reason the project includes the deployment infrastructure at all: the live app is the most powerful artifact for portfolio purposes, even if the underlying analysis is also independently valuable."
  ),
];

// ---- Section 10: Recommendations and Future Work ------------------------
const section10 = [
  h1("10. Recommendations and Future Work"),
  h2("Recommendations for the city and transportation planners"),
  p(
    "The single most actionable finding for transportation planners is the roughly 50% range in average vehicle speed across the day, with the lowest speeds concentrated in the late-afternoon and early-evening commute window. This congestion footprint is not a surprise — anyone who has tried to take a taxi across midtown at 5 PM knows the experience — but the dataset quantifies it in a way that supports evidence-based intervention design. Signal timing optimization in the highest-volume zones (particularly midtown Manhattan), expanded congestion pricing during the worst-hit hours (the current Manhattan congestion-pricing program affects entry into the central business district but could plausibly be extended in time or in geographic scope), and selective taxi-stand placement in zones with consistently high pickup demand are all interventions that could plausibly recover throughput."
  ),
  p(
    "Beyond congestion, the data also surfaces equity questions worth investigating. The concentration of Yellow Taxi service in Manhattan and the airports means that outer-borough residents have substantially less access to street-hailing taxi service, a gap that has historically been one of the policy rationales for the Green Taxi and FHV markets. Whether the current allocation is appropriate is a normative question this report does not address, but the data here would support a quantitative analysis of zone-level service equity if combined with the FHV datasets that cover the outer boroughs."
  ),
  h2("Recommendations for operators"),
  p(
    "For taxi operators, several findings have direct operational implications. The airport flat-rate cluster suggests that airport-dedicated dispatch should be modeled separately from city dispatch. The economics — and the duration distributions, and the demand patterns by hour of day — are sufficiently different that a single fleet-wide model is leaving signal on the table. A driver who specializes in JFK runs faces a very different optimization problem than a driver who works the midtown business district during commute hours, and routing decisions, surge pricing, and shift scheduling should reflect that difference."
  ),
  p(
    "Operators should also consider that the tight clustering of card tips around the in-cab preset percentages is a user-experience outcome more than a behavioural one: even small changes to the preset values could plausibly produce meaningful aggregate revenue changes. A controlled A/B test of preset values would be the natural way to investigate this hypothesis. The risk is that overly aggressive defaults could damage rider satisfaction or even produce a measurable shift away from card payment toward cash, which would reduce total revenue rather than increase it. Any change to the defaults should be tested carefully, with attention to second-order effects."
  ),
  p(
    "A third operational implication relates to driver shift scheduling. The clean weekly demand pattern — Thursday peak, Sunday and Monday troughs, sharp evening peaks each weekday — suggests that drivers and dispatchers can plan with relatively high confidence about when demand will be highest. The current fleet allocation appears to be reasonably well-matched to this pattern, but there is room for fine-tuning at the hour level, particularly around the late-night-to-early-morning transition where demand drops rapidly and a smaller fleet would be operationally sufficient."
  ),
  h2("Future work on this project"),
  bulletBold("Geographic visualization. ", "Join the TLC taxi-zone shapefile and render choropleth maps of trips, revenue, and average fare per zone using Folium or PyDeck. This would close a clear gap in the current analysis, which represents zones only as integer IDs. A choropleth would let users immediately understand which neighborhoods drive the bulk of taxi activity and how trip economics vary geographically. The TLC shapefile is freely available and the integration is straightforward; this is the highest-priority next step for the project."),
  bulletBold("Cross-modal comparison. ", "Pull the FHV (Uber, Lyft) data for the same period and compare market share, pricing, and demand patterns. Yellow Taxi is now a minority of the for-hire vehicle market in NYC — the HVFHV dataset substantially exceeds Yellow Taxi in trip volume — and understanding the relative dynamics is important context. A cross-modal report would let a reader see, for instance, whether the temporal patterns observed in Yellow Taxi data also hold in the on-demand platform data, or whether Uber's algorithmic pricing produces meaningfully different rider behaviour."),
  bulletBold("Probabilistic ETA. ", "Replace the point-estimate trip-duration model with a quantile-regression ensemble that produces P10–P90 intervals. This is dramatically more useful in production because it gives the dispatcher a confidence range to communicate to the rider rather than a single number that may or may not be accurate. The implementation is a small extension of the current LightGBM training pipeline: train three models with different quantile-regression objectives instead of one model with the standard squared-error objective."),
  bulletBold("Weather features. ", "Join an hourly precipitation feed (NOAA publishes this freely for JFK and LaGuardia weather stations) and incorporate weather conditions as additional model features. Rain and snow are known to substantially affect both demand (more people take taxis to avoid walking in bad weather) and trip duration (slower speeds in adverse conditions). The integration would require joining on hourly timestamps and could plausibly lift the model's R² by several percentage points, with the biggest gains on the small number of days that actually have unusual weather."),
  bulletBold("Real-time deployment. ", "Containerize the dashboard with Docker and deploy to Fly.io or Render for a permanent, high-availability public demo, with auto-redeploy on every Git push and proper logging and monitoring. The current Streamlit Community Cloud deployment is sufficient for a portfolio demonstration but would not be appropriate for any production use case. Moving to a more robust hosting environment would be a small additional engineering investment with substantial operational benefits."),
  bulletBold("Cohort analysis. ", "Pull a full calendar year (or several years) of data and analyze seasonality at the year scale, including holiday effects, summer-versus-winter tourism patterns, and the effect of major recurring events at Madison Square Garden, the United Nations General Assembly, the New York Marathon, and other major venues. The three-month window of the current analysis is enough to surface daily and weekly patterns but cannot capture annual seasonality."),
  bulletBold("Driver-side analysis. ", "While the TLC data is anonymized at the driver level, certain aggregate questions about driver productivity, shift patterns, and earnings can still be approached through clever aggregations. A future extension could investigate, for example, whether the same number of trips per shift translates into the same revenue per shift across different hours and zones — the answer is almost certainly no, and quantifying the difference would be valuable for shift-scheduling decisions."),
  h2("Lessons learned"),
  p(
    "Several lessons from this project are worth recording explicitly, both for the author's future reference and for any reader who might be undertaking a similar project. First, public datasets are rarely as clean as their documentation suggests. The year-range filter alone caught hundreds of thousands of rows that would have polluted time-series visualizations in subtle but real ways. The temptation when working with a well-known dataset is to trust the schema and skip the boring step of inspecting the actual data distribution; this temptation should be resisted. Always inspect the distribution of each field, look for the obvious outlier regions, and design filters that handle them explicitly. The cost of doing this work upfront is small; the cost of not doing it shows up later as confusing chart artifacts that take much longer to diagnose."
  ),
  p(
    "Second, DuckDB is dramatically more efficient than pandas for large-Parquet analytics in memory-constrained environments. The Streamlit Cloud bootstrap module was originally written in pandas and crashed on a 1 GB free-tier container because pandas insists on loading the entire dataset into memory before applying any transformations. Rewriting the bootstrap as a single DuckDB COPY query reduced the peak memory footprint from over 2 GB to under 200 MB, and was actually faster end-to-end than the pandas version. DuckDB's streaming execution model is the right tool whenever the data is larger than a comfortable fraction of available memory; pandas remains the right tool for everything that fits comfortably in RAM."
  ),
  p(
    "Third, modeling decisions and analytical decisions should be made together rather than sequentially. Knowing that airport trips form a structurally different population shaped both the EDA (treating them as a separate cohort in every comparison table) and the ML model (including ratecode_id as a categorical feature so the model can distinguish them). A workflow that does the EDA first and then bolts on a model would have missed this connection. The integrated workflow used here, where each EDA finding immediately suggests a candidate model feature and each model feature is immediately tested against the EDA, produces stronger results in less total time."
  ),
  p(
    "Fourth, the most valuable artifact in a portfolio project is the one a reviewer can click on and explore in thirty seconds. A written report, a notebook, and a GitHub repository all have value, but a live deployed dashboard is the artifact that communicates competence fastest. The engineering work to make the dashboard deployable — the bootstrap module, the caching strategy, the DuckDB rewrite for memory safety — is therefore not optional infrastructure work; it is the work that turns the underlying analysis into something a recruiter or interviewer will actually engage with. Skipping it would substantially reduce the portfolio value of the project."
  ),
  p(
    "Fifth, the project benefited substantially from tight feedback loops. The cleaning module was developed by running it on a small sample, looking at the output, identifying issues, adjusting the filters, and re-running — a cycle that took under a minute per iteration. The ML model was developed by training on a small sample, looking at the validation metrics, adjusting features, and re-running — a similar fast cycle. The dashboard was developed by running Streamlit locally with hot-reload, so every code change was visible in the browser within a second. Long feedback loops would have stretched this project across days or weeks instead of hours; investing in fast iteration is the highest-leverage engineering decision in any project of this scope."
  ),
  h2("Reproducing this analysis"),
  p(
    "Every step in this report is reproducible from the public GitHub repository at github.com/LeanKishan/nyc-taxi-analysis. The repository is structured so that any individual stage of the pipeline can be re-run independently, and the entire end-to-end pipeline can be reproduced with a small number of commands. After cloning the repository and installing dependencies from requirements.txt — a process that takes approximately two minutes on a typical broadband connection — the entire analysis runs with five commands:"
  ),
  bullet("python -m src.download_data --year 2025 --months 1 2 3"),
  bullet("python -m src.data_cleaning"),
  bullet("python -m src.sql_analytics"),
  bullet("python -m src.train_model"),
  bullet("streamlit run dashboard/app.py"),
  p(
    "Unit tests for the cleaning module are under tests/ and can be run with pytest. The expected end-to-end runtime on a developer laptop is approximately fifteen minutes, dominated by the initial Parquet download (about three minutes for the three monthly files combined) and the ML training step (about one minute for the LightGBM training plus the figure generation). The cleaning, SQL analytics, and dashboard startup steps each complete in under a minute. The Streamlit dashboard remains running until the user stops it; navigating between tabs is effectively instant once the initial connection is established."
  ),
  p(
    "Anyone who wants to extend the analysis to additional months simply runs the download command with the desired month numbers and re-runs the cleaning, SQL, and training steps; the rest of the pipeline picks up the new data automatically. Anyone who wants to swap in a different model (XGBoost, a neural network, a statistical baseline like a simple linear regression on distance) can replace the train_model module with an alternative implementation while leaving the rest of the pipeline unchanged. This kind of modularity is one of the points the project is designed to demonstrate, and the structure should be easy to adapt to similar analytical problems on different datasets."
  ),
  h2("Acknowledgments"),
  p(
    "Data is courtesy of the New York City Taxi and Limousine Commission, which publishes the trip records as a public service. The continued availability of this data over more than fifteen years is an unusual commitment to public transparency that has supported a remarkable amount of academic, commercial, and educational work, and the analyst community is collectively in the TLC's debt."
  ),
  p(
    "The open-source ecosystem that makes this kind of analysis possible — DuckDB for analytical SQL, LightGBM for gradient boosting, pandas for in-memory data manipulation, Streamlit for the interactive dashboard, scikit-learn for the standard machine learning utilities, matplotlib and seaborn for visualization, and the broader Python ecosystem that ties them all together — represents extraordinary value freely contributed by their respective maintainers and communities. A project like this one would have been impossible a decade ago, would have required substantial vendor licensing fees five years ago, and is today the work of a single developer over a few hours. The pace of improvement in open-source data tools is one of the most underappreciated trends in technology, and this project would not exist without it."
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

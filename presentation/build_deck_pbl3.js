// PBL 3 complete presentation (notebook walkthrough layout). Run: node build_deck_pbl3.js
// Needs presentation/assets/boosting_errors.png  (python presentation/make_deck_figures.py)
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const ROOT = path.join(__dirname, "..");
const UNIT_OF = { "01_dataset": "Unit_1_2_PBL1_Data_and_Regression", "05_regression": "Unit_1_2_PBL1_Data_and_Regression",
  "02_clustering": "Unit_3_PBL2_Clustering_KMeans", "03_hmm": "Unit_4_PBL2_HMM", "04_comparison": "Unit_4_PBL2_HMM",
  "06_boosting": "Unit_5_PBL3_Gradient_Boosting" };
const RES = (sub, ...p) => path.join(ROOT, UNIT_OF[sub], "results", sub, ...p);
const OUT = path.join(__dirname, "AI_Supply_Chain_Optimizer_PBL3_Complete_Presentation.pptx");

const THEME = {
  name: "Academic Slate", headFontFace: "Calibri", bodyFontFace: "Calibri",
  colors: { dk1: "0F172A", lt1: "FFFFFF", dk2: "111827", lt2: "F8FAFC", accent1: "2563EB", accent2: "0891B2",
            accent3: "EB6834", accent4: "1BAF7A", accent5: "64748B", accent6: "334155", hlink: "2563EB", folHlink: "64748B" },
};
const HEX = THEME.colors;

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "AI Supply Chain Optimizer – PBL 3: Ensemble Learning and Boosting";
  pres.author = "Project team";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const X = 0.8, CW = 11.73;

  pres.defineSlideMaster({ title: "TITLE", background: { color: C.background1 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1.0, y: 2.25, w: 11.33, h: 1.6, fontSize: 34, bold: true,
        color: C.text2, valign: "top", align: "left", margin: 0 }, text: "Title" } }] });
  pres.defineSlideMaster({ title: "CONTENT", background: { color: C.background1 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: X, y: 0.4, w: CW, h: 0.8, fontSize: 26, bold: true,
        color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "Title" } }] });

  const content = (title, section) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontSize: 14, color: C.accent6, valign: "top", margin: 0 }, o));
  const bullets = (s, items, o) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 7 } })),
    Object.assign({ isTextBox: true, fontSize: 14, color: C.accent6, valign: "top", margin: 0.02 }, o));
  const heading = (s, t, x, y, w, color) => txt(s, t, { x, y, w, h: 0.4, fontSize: 16, bold: true, color: color || C.accent1 });
  const panel = (s, x, y, w, h, name) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background2 }, line: { color: "E2E8F0", width: 0.75 }, objectName: name });
  const imgW = (s, file, x, y, w, ratio, name) => s.addImage({ path: file, x, y, w, h: w / ratio, objectName: name });
  // table helpers
  const H = (t, o) => ({ text: t, options: Object.assign({ bold: true, color: HEX.lt1, fill: { color: HEX.dk1 }, fontSize: 12, valign: "middle" }, o || {}) });
  const D = (t, r, o) => ({ text: t, options: Object.assign({ color: HEX.accent6, fill: { color: r % 2 ? HEX.lt1 : HEX.lt2 }, fontSize: 12, valign: "middle" }, o || {}) });
  const tbl = (s, rows, o) => s.addTable(rows, Object.assign({ x: X, w: CW, border: { type: "solid", pt: 0.75, color: "E2E8F0" }, margin: 0.06 }, o));
  const sub = (t) => ({ text: t, options: { subscript: true } });

  // ================================================================ 1 title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Introduction" });
  txt(s, "PBL-3: Ensemble Learning and Boosting", { x: 1.0, y: 1.65, w: 11.33, h: 0.5, fontSize: 22, bold: true, color: C.accent1 });
  s.addText("AI Supply Chain Optimizer: Forecasting Weekly Quick-Commerce Food Demand with Gradient Boosting", { placeholder: "title" });
  txt(s, "B.Tech Machine Learning Review  |  21CSC305P Machine Learning", { x: 1.0, y: 4.05, w: 11.33, h: 0.4, fontSize: 16, color: C.accent5 });
  txt(s, "Dataset: India Quick Commerce Orders 2025–2026 (Kaggle, 58 dark stores, 1,95,148 orders)", { x: 1.0, y: 4.5, w: 11.33, h: 0.4, fontSize: 15, bold: true, color: C.accent2 });
  txt(s, "Team: [Name – Reg No]  ·  Faculty: [Name]  ·  November 2026", { x: 1.0, y: 5.5, w: 11.33, h: 0.4, fontSize: 14, color: C.accent5 });
  s.addNotes("This review covers the ensemble learning part of the project: gradient boosting applied to the same weekly demand forecasting task as the regression models, with the same data split and metrics.");

  // ================================================================ 2 dataset
  pres.addSection({ title: "Data" });
  s = content("Dataset: Weekly Food Orders of Quick-Commerce Dark Stores", "Data");
  txt(s, "Loaded store-week modelling table (store_weekly.csv): (4,039, 13)  —  df.head()", { x: X, y: 1.3, w: CW, h: 0.4, fontSize: 13, bold: true, color: C.text2 });
  const headCols = ["store_id", "week", "orders", "lag1", "lag2", "lag4", "roll4", "store_age_w", "festival_days", "monsoon", "area_Premium Residential", "area_Residential", "area_Student / PG Hub"];
  const headRows = [["DS001", "03-02-2025", 20, "20.0", "17.0", "12.0", "17.50", 4, 0, 0, 0, 0, 1],
                    ["DS001", "10-02-2025", 26, "20.0", "20.0", "21.0", "19.50", 5, 2, 0, 0, 0, 1],
                    ["DS001", "17-02-2025", 19, "26.0", "20.0", "17.0", "20.75", 6, 0, 0, 0, 0, 1],
                    ["DS001", "24-02-2025", 25, "19.0", "26.0", "20.0", "21.25", 7, 0, 0, 0, 0, 1],
                    ["DS001", "03-03-2025", 26, "25.0", "19.0", "20.0", "22.50", 8, 0, 0, 0, 0, 1]];
  tbl(s, [headCols.map((c) => H(c, { fontSize: 11, align: "center" })),
          ...headRows.map((r, i) => r.map((v) => D(String(v), i, { fontSize: 11, align: "center" })))],
      { y: 1.75, colW: [0.8, 1.05, 0.7, 0.65, 0.65, 0.65, 0.7, 1.0, 1.05, 0.85, 1.25, 1.18, 1.2], rowH: [0.62, 0.34, 0.34, 0.34, 0.34, 0.34], objectName: "df head table" });
  bullets(s, [
    "Source: Kaggle “India Quick Commerce Orders 2025–2026” (CC0) — 1,95,148 orders × 39 columns from 58 dark stores (small delivery-only warehouses).",
    "Modelling table: 4,039 store-weeks (58 stores × up to 86 weeks, 03-02-2025 to 21-09-2026). Target = number of food orders a store receives in the week.",
    "Features are known before the week starts: orders 1, 2 and 4 weeks earlier (lag1, lag2, lag4), the 4-week average (roll4), store age, festival days, monsoon flag and location type (one-hot area_* columns).",
    "Fits the task: a real forecasting problem with steady growth, festival spikes and stores of very different sizes (1 to 92 orders a week) — a good test for a flexible ensemble.",
  ], { x: X, y: 4.35, w: CW, h: 2.75 });
  s.addNotes("The same table is used for the regression models (PBL 1), so every model is compared on identical rows.");

  // ================================================================ 3 quality check
  s = content("Data Loading and Quality Check", "Data");
  panel(s, X, 1.3, 6.55, 4.75, "df info panel");
  txt(s, [
    "<class 'pandas.DataFrame'>",
    "RangeIndex: 4039 entries, 0 to 4038",
    "Data columns (total 13 columns):",
    " #   Column                    Non-Null Count  Dtype",
    "---  ------                    --------------  -----",
    " 0   store_id                  4039 non-null   str",
    " 1   week                      4039 non-null   datetime64",
    " 2   orders                    4039 non-null   int64",
    " 3   lag1                      4039 non-null   float64",
    " 4   lag2                      4039 non-null   float64",
    " 5   lag4                      4039 non-null   float64",
    " 6   roll4                     4039 non-null   float64",
    " 7   store_age_w               4039 non-null   int64",
    " 8   festival_days             4039 non-null   int64",
    " 9   monsoon                   4039 non-null   int64",
    " 10  area_Premium Residential  4039 non-null   int64",
    " 11  area_Residential          4039 non-null   int64",
    " 12  area_Student / PG Hub     4039 non-null   int64",
    "dtypes: datetime64(1), float64(4), int64(7), str(1)",
    "memory usage: 410.3 KB",
    "Missing values: 0 | Duplicates: 0",
  ].join("\n"), { x: X + 0.2, y: 1.5, w: 6.25, h: 5.45, fontSize: 12, fontFace: "Courier New", color: C.text1 });
  bullets(s, [
    "Complete data: 0 missing values and 0 duplicate rows in all 13 columns, so no imputation was needed.",
    "Types: target and features are numeric (int64 / float64); store_id is text and week is a date.",
    "Light-weight: 4,039 rows use about 410 KB of memory, so the full grid search runs in under a minute.",
    "Cleaned upstream (notebook 01): only food categories kept (1,17,077 orders) and the IPL-evening surge removed (1,13,390 orders).",
    "The first 4 weeks of each store are dropped, because lag4 and roll4 need 4 weeks of history.",
    "Time-based split, never random: train = 60 weeks up to 23-03-2026 (2,531 rows); test = 26 later weeks (1,508 rows). Shuffling would leak the future into training.",
  ], { x: 7.65, y: 1.3, w: 4.88, h: 5.75 });
  s.addNotes("Output of df.info() on store_weekly.csv. The split date 30-03-2026 is the same as in the regression notebook.");

  // ================================================================ 4 preprocessing
  s = content("Preprocessing Pipeline", "Data");
  txt(s, "1,95,148 raw orders  →  Food filter + IPL removal  →  Weekly totals per store  →  Lag & rolling features  →  Time split  →  Growth-free target",
    { x: X, y: 1.3, w: CW, h: 0.5, fontSize: 14, bold: true, color: C.accent1, align: "center", valign: "middle" });
  const pp = [["Cleaning", "Food categories kept and the IPL match-evening surge removed: 1,95,148 → 1,13,390 orders."],
              ["Weekly aggregation", "Orders counted per store per week (weeks start on Monday): 58 stores, 4,039 usable store-weeks."],
              ["Lag features", "lag1, lag2, lag4 and roll4 are built only from past weeks (shifted), so nothing from the forecast week leaks in."],
              ["Time split", "Train 03-02-2025 – 23-03-2026 (2,531 rows); test 30-03-2026 – 21-09-2026 (1,508 rows) — the same split as PBL 1."],
              ["Growth-free target", "z = log(orders ÷ roll4); relative lags r1, r2, r4 = lag ÷ roll4; woy = week of the year. Forecast = roll4 × e^ẑ."]];
  tbl(s, [[H("Step", { fontSize: 13 }), H("Execution Details", { fontSize: 13 })], ...pp.map((r, i) => [D(r[0], i, { bold: true, fontSize: 13, color: HEX.dk2 }), D(r[1], i, { fontSize: 13 })])],
      { y: 2.0, colW: [2.6, 9.13], rowH: [0.5, 0.62, 0.62, 0.62, 0.62, 0.62], objectName: "pipeline table" });
  txt(s, "Trees need no feature scaling, so StandardScaler is used only for the regression baselines. Every model is scored with MAE, RMSE, R² and WAPE on the same 1,508 test store-weeks; MAE (average error in orders) is the main metric.",
    { x: X, y: 5.85, w: CW, h: 0.9, fontSize: 13, color: C.accent5, align: "center" });
  s.addNotes("The growth-free target is explained on the feature engineering slide.");

  // ================================================================ 5 EDA
  s = content("Exploratory Data Analysis: Demand Distribution, Correlations and Trend", "Data");
  imgW(s, RES("05_regression", "eda_weekly.png"), X, 1.3, 7.75, 1546 / 1090, "eda charts");
  bullets(s, [
    "Weekly orders are right-skewed: median 25, mean 26.9, range 1 to 92 per store-week.",
    "Past demand is the strongest signal: roll4 correlates 0.88 with orders and lag1, lag2, lag4 about 0.81.",
    "Festival days have a weak linear correlation (≈ 0.1) yet cause sharp spikes (Oct–Nov 2025, Aug–Sep 2026) — a non-linear effect that trees can capture.",
    "Demand roughly doubles over the period in every location type. Steady growth like this is hard for tree models (next slide).",
  ], { x: 8.8, y: 1.35, w: 3.73, h: 5.7, fontSize: 15 });
  s.addNotes("Top left: histogram of weekly orders. Top right: Pearson correlation with weekly orders. Bottom: average weekly orders per store by location type.");

  // ================================================================ 6 feature engineering
  pres.addSection({ title: "Method" });
  s = content("Feature Engineering: A Growth-Free Target for Trees", "Method");
  imgW(s, RES("06_boosting", "boosting_feature_analysis.png"), X, 1.3, 7.5, 1546 / 658, "extrapolation charts");
  panel(s, X, 4.75, 7.5, 2.3, "takeaway panel");
  heading(s, "Key idea", X + 0.25, 4.9, 7.0);
  txt(s, "Split each forecast into two parts. The recent level roll4 carries the store's size and growth; the trees only learn the change from that level, caused by festivals, season and momentum. Forecast = roll4 × e^ẑ.",
    { x: X + 0.25, y: 5.35, w: 7.0, h: 1.6, fontSize: 15 });
  bullets(s, [
    "Left plot: most stores' average test week is close to or above their best training week; 19.2% of test store-weeks exceed the store's training maximum.",
    "A regression tree predicts averages of training values, so it can never forecast above what it has seen. Plain boosting on order counts scored MAE 5.92 — worse than linear regression (5.35).",
    "Fix: target z = log(orders ÷ roll4) with relative lags r1, r2, r4 = lag ÷ roll4 and week of year (woy).",
    "Right plot: the new target is stable and nearly symmetric (mean +0.011, std 0.270), with the same range in training and test.",
  ], { x: 8.6, y: 1.3, w: 3.93, h: 5.75 });
  s.addNotes("This reformulation is the key design decision of this part of the project.");

  // ================================================================ 7 how GB works
  s = content("How Gradient Boosting Works: Step by Step", "Method");
  const gs = [["1. Initial model F₀", "Predict one constant for every row: the mean of the target.", "F₀ = 0.013, so the first forecast ≈ roll4 × 1.013."],
              ["2. Pseudo-residuals", "rᵢ = yᵢ − Fₘ₋₁(xᵢ): the negative gradient of the squared-error loss, i.e. what is still wrong.", "How far each store-week is from its current forecast."],
              ["3. Fit a weak learner", "A small regression tree hₘ is trained to predict the residuals.", "Depth-2 tree on a random 80% of rows (subsample = 0.8)."],
              ["4. Shrinkage update", "Fₘ = Fₘ₋₁ + ν·hₘ — add only a fraction ν of the tree's correction.", "ν = 0.02: each tree fixes only 2% of what it finds."],
              ["5. Repeat", "Repeat steps 2–4 for m = 1 … M; many small corrections add up.", "M = 200 trees, chosen by time-series cross-validation."]];
  tbl(s, [[H("Step", { fontSize: 13 }), H("What happens", { fontSize: 13 }), H("In our project", { fontSize: 13 })],
          ...gs.map((r, i) => [D(r[0], i, { bold: true, fontSize: 13, color: HEX.dk2 }), D(r[1], i, { fontSize: 13 }), D(r[2], i, { fontSize: 13 })])],
      { y: 1.3, colW: [2.6, 5.0, 4.13], rowH: [0.45, 0.62, 0.62, 0.62, 0.62, 0.62], objectName: "steps table" });
  panel(s, X, 5.2, CW, 1.85, "formula panel");
  heading(s, "Final model", X + 0.25, 5.32, 3.0);
  txt(s, [{ text: "F" }, sub("M"), { text: "(x) = F" }, sub("0"), { text: " + ν·h" }, sub("1"), { text: "(x) + ν·h" }, sub("2"), { text: "(x) + … + ν·h" }, sub("M"), { text: "(x)" }],
    { x: X + 0.25, y: 5.72, w: 6.5, h: 0.5, fontSize: 22, bold: true, color: C.text2 });
  txt(s, [{ text: "Example from our model: ", options: { bold: true, color: HEX.accent1 } },
          { text: "the first tree splits on festival_days ≤ 1.5 and then week of year ≤ 37, so festival weeks after mid-September get the largest upward correction (+0.385 before shrinkage)." }],
    { x: X + 0.25, y: 6.25, w: 11.2, h: 0.75, fontSize: 13 });
  s.addNotes("Gradient boosting is gradient descent in function space. With squared-error loss the negative gradient equals the residual, so every tree is trained on the current errors.");

  // ================================================================ 8 models compared
  s = content("Models Compared: Methodologies and Hyperparameters", "Method");
  const mc = [["Naive (last week)", "Baseline", "Forecast = orders in the previous week (lag1).", "none"],
              ["Linear Regression", "Linear model (PBL 1)", "Least-squares weighted sum of the 10 standardised features.", "none"],
              ["Ridge Regression", "Regularised linear (PBL 1)", "Linear regression with an L2 penalty that shrinks the weights.", "α by time-series CV (0.001)"],
              ["Bayesian Regression", "Probabilistic linear (PBL 1)", "Prior on the weights; gives a mean forecast and a 95% interval.", "priors learned from data"],
              ["Gradient Boosting (original)", "Boosting ensemble", "Trees fitted one after another to the residuals of order counts.", "200 trees, depth 2, ν = 0.05, subsample 0.8"],
              ["Gradient Boosting (growth-free)", "Boosting ensemble", "Same algorithm on z = log(orders ÷ roll4) with relative features.", "200 trees, depth 2, ν = 0.02, subsample 0.8"]];
  tbl(s, [[H("Model"), H("Type"), H("How it Works"), H("Hyperparameters")],
          ...mc.map((r, i) => [D(r[0], i, { bold: true, color: HEX.dk2 }), D(r[1], i), D(r[2], i), D(r[3], i)])],
      { y: 1.3, colW: [2.75, 2.35, 4.08, 2.55], rowH: [0.42, 0.55, 0.55, 0.55, 0.55, 0.62, 0.62], objectName: "models table" });
  heading(s, "Conceptual Contrast: Linear Regression vs. Gradient Boosting", X, 5.35, CW);
  bullets(s, [
    "Linear regression adds a fixed amount per feature and can extend a growth trend, but it cannot learn interactions such as festival × season by itself.",
    "Gradient boosting learns non-linear effects and interactions automatically, but only inside the range of values seen in training — hence the growth-free target.",
  ], { x: X, y: 5.8, w: CW, h: 1.25 });
  s.addNotes("All six models use the same training weeks, test weeks and metrics.");

  // ================================================================ 9 tuning
  s = content("Hyperparameter Tuning: Time-Series Cross-Validation", "Method");
  const cv = [[200, 0.02, 2, 0.2167], [200, 0.02, 3, 0.2189], [400, 0.02, 2, 0.2190], [200, 0.05, 2, 0.2212], [400, 0.02, 3, 0.2227], [800, 0.02, 2, 0.2233],
              [200, 0.05, 3, 0.2249], [400, 0.05, 2, 0.2260], [800, 0.02, 3, 0.2285], [800, 0.05, 2, 0.2326], [400, 0.05, 3, 0.2332], [800, 0.05, 3, 0.2409]];
  tbl(s, [["Rank", "n_estimators", "learning_rate", "max_depth", "CV MAE"].map((t) => H(t, { align: "center" })),
          ...cv.map((r, i) => {
            const o = i === 0 ? { bold: true, color: HEX.accent1, fill: { color: "DBEAFE" }, align: "center" } : { align: "center" };
            return [D(String(i + 1), i, o), D(String(r[0]), i, o), D(String(r[1]), i, o), D(String(r[2]), i, o), D(r[3].toFixed(4), i, o)];
          })],
      { x: X, w: 6.4, y: 1.3, colW: [0.8, 1.45, 1.5, 1.3, 1.35], rowH: 0.4, objectName: "cv table" });
  txt(s, "CV MAE is measured on the log-ratio target (lower is better).", { x: X, y: 6.62, w: 6.4, h: 0.35, fontSize: 12, color: C.accent5, italic: true });
  bullets(s, [
    "GridSearchCV tried 3 × 2 × 2 = 12 combinations: n_estimators {200, 400, 800}, learning_rate {0.02, 0.05}, max_depth {2, 3}. Search time: 36 s.",
    "TimeSeriesSplit(5): each fold trains on earlier weeks and validates on the block that follows, so the future is never used to choose settings.",
    "Every step towards a bigger model — more trees, a higher learning rate or deeper trees — raised the validation error: a clear sign of over-fitting.",
    "Selected: the smallest model, 200 trees of depth 2 with learning rate 0.02 (CV MAE 0.2167, worst 0.2409).",
  ], { x: 7.6, y: 1.3, w: 4.93, h: 5.75, fontSize: 15 });
  s.addNotes("All 12 combinations, ranked by cross-validated MAE.");

  // ================================================================ 10 results
  pres.addSection({ title: "Results" });
  s = content("Results: All Models on the Same 26 Test Weeks", "Results");
  const res = [["Naive (last week)", "Baseline", 6.761, 8.918, 0.644, "20.7%"], ["Linear Regression", "PBL 1", 5.353, 6.970, 0.782, "16.4%"],
               ["Ridge Regression", "PBL 1", 5.353, 6.970, 0.782, "16.4%"], ["Bayesian Regression", "PBL 1", 5.356, 6.975, 0.782, "16.4%"],
               ["Gradient Boosting (original features)", "PBL 3", 5.917, 8.050, 0.710, "18.1%"], ["Gradient Boosting (growth-free target)", "PBL 3", 5.315, 6.939, 0.784, "16.3%"]];
  tbl(s, [["Model", "Part", "MAE", "RMSE", "R²", "WAPE"].map((t, j) => H(t, { fontSize: 13, align: j ? "center" : "left" })),
          ...res.map((r, i) => {
            const best = i === res.length - 1;
            const o = best ? { bold: true, color: HEX.accent1, fill: { color: "DBEAFE" }, fontSize: 13 } : { fontSize: 13 };
            return [D(r[0], i, Object.assign({}, o, best ? {} : { bold: true, color: HEX.dk2 })), D(r[1], i, Object.assign({ align: "center" }, o)),
                    D(r[2].toFixed(3), i, Object.assign({ align: "center" }, o)), D(r[3].toFixed(3), i, Object.assign({ align: "center" }, o)),
                    D(r[4].toFixed(3), i, Object.assign({ align: "center" }, o)), D(r[5], i, Object.assign({ align: "center" }, o))];
          })],
      { y: 1.3, colW: [4.4, 1.3, 1.5, 1.5, 1.5, 1.53], rowH: 0.48, objectName: "results table" });
  bullets(s, [
    "Best model: gradient boosting with the growth-free target — MAE 5.315, RMSE 6.939, R² 0.784 and WAPE 16.3%, the best value on all four metrics.",
    "Ranking by MAE: Boosting (growth-free) 5.315 < Linear = Ridge 5.353 < Bayesian 5.356 < Boosting (original) 5.917 < Naive 6.761.",
    "The growth-free target cut boosting's error by 10% (MAE 5.92 → 5.32) and RMSE by 14% (8.05 → 6.94).",
    "All good models sit about one order above the noise floor (MAE ≈ 4.34): most of the remaining error is random week-to-week variation.",
  ], { x: X, y: 4.85, w: CW, h: 2.2, fontSize: 15 });
  s.addNotes("MAE = average absolute error in orders per store-week. RMSE penalises large errors more. R² = share of variation explained. WAPE = total absolute error as a share of total orders.");

  // ================================================================ 11 importance + learning curve
  s = content("Interpretation: Feature Importance and Boosting Rounds", "Results");
  imgW(s, RES("06_boosting", "importance_and_learning_curve.png"), (13.333 - 9.8) / 2, 1.3, 9.8, 1570 / 658, "importance and learning curve");
  bullets(s, [
    "Left (feature importance): festival_days (0.34) and roll4 (0.21) dominate, followed by r4 (0.17) and week of year (0.12). A permutation test on the test weeks confirms festival_days and roll4 as the features that clearly matter.",
    "Right (learning curve): training error keeps falling, but validation error is lowest at 133 trees (MAE 4.70) and rises to 4.87 at 800 trees — over-fitting. The chosen 200 trees with ν = 0.02 stay close to the minimum.",
  ], { x: X, y: 5.55, w: CW, h: 1.5 });
  s.addNotes("Learning curve from staged_predict; validation = the last 13 training weeks.");

  // ================================================================ 12 errors
  s = content("Interpretation: Forecast Accuracy and Errors", "Results");
  imgW(s, path.join(__dirname, "assets", "boosting_errors.png"), (13.333 - 9.6) / 2, 1.3, 9.6, 1584 / 638, "errors charts");
  bullets(s, [
    "Left: forecasts follow the diagonal; 57% of store-week forecasts are within ±5 orders. Festival weeks (orange) are the hardest to predict.",
    "Boosting's gain comes from festival weeks: MAE 6.49 vs 7.30 for linear regression (−11%). On normal weeks the two are equal (5.27 vs 5.28).",
    "Right: all models track the network total but lag the steady growth (about 1.7 orders under per store-week). Plain boosting (dashed) falls furthest behind after August because its trees cannot exceed training values.",
  ], { x: X, y: 5.3, w: CW, h: 1.75 });
  s.addNotes("Computed from the saved test predictions of notebooks 05 and 06; 58 of the 1,508 test store-weeks contain festival days.");

  // ================================================================ 13 findings
  s = content("Comparison and Key Findings", "Results");
  const cmp = [["Naive (last week)", "6.761", "0.644", "+26.3%"], ["Linear Regression (PBL 1)", "5.353", "0.782", "—"], ["Bayesian Regression (PBL 1)", "5.356", "0.782", "+0.1%"],
               ["Gradient Boosting (original)", "5.917", "0.710", "+10.5%"], ["Gradient Boosting (growth-free)", "5.315", "0.784", "−0.7%"]];
  tbl(s, [[H("Model"), H("MAE", { align: "center" }), H("R²", { align: "center" }), H("MAE vs. linear", { align: "center" })],
          ...cmp.map((r, i) => {
            const best = i === cmp.length - 1, o = best ? { bold: true, color: HEX.accent1, fill: { color: "DBEAFE" } } : {};
            return [D(r[0], i, Object.assign({ bold: true, color: HEX.dk2 }, o)), D(r[1], i, Object.assign({ align: "center" }, o)), D(r[2], i, Object.assign({ align: "center" }, o)), D(r[3], i, Object.assign({ align: "center" }, o))];
          })],
      { x: X, w: 7.5, y: 1.3, colW: [3.3, 1.3, 1.3, 1.6], rowH: 0.5, objectName: "comparison table" });
  panel(s, X, 4.35, 7.5, 2.7, "connection panel");
  heading(s, "Connection to the Broader Project (AI Supply Chain Optimizer)", X + 0.25, 4.47, 7.0, C.accent2);
  txt(s, "PBL 1 regression gives a transparent baseline forecast with a 95% interval. PBL 2 groups the stores into 3 types (K-Means) and detects Normal, Monsoon and Festival-peak demand regimes (HMM). PBL 3 boosting adds the most accurate weekly forecast, especially in festival weeks. Together they tell each dark store how much food to stock and when demand will spike.",
    { x: X + 0.25, y: 4.92, w: 7.0, h: 2.05, fontSize: 14.5 });
  heading(s, "Final Takeaways", 8.6, 1.3, 3.93);
  bullets(s, [
    "Boosting gives the best forecast, but only after the growth-free target: MAE 5.32 vs 5.35 for linear and 5.92 for plain boosting.",
    "Problem design mattered more than the algorithm: the same model went from worse than linear to the best model.",
    "Small, heavily shrunk models generalise best: 200 depth-2 trees with learning rate 0.02.",
    "Festival days and the recent level drive demand. The error is near the noise floor, so further gains need new data (prices, promotions, weather).",
  ], { x: 8.6, y: 1.75, w: 3.93, h: 5.3, fontSize: 15 });
  s.addNotes("Gradient boosting confirms that the regression baseline was already close to the limit of the available features.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

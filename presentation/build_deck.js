// Final project deck: AI Supply Chain Optimizer (PBL 1–3). Run: node build_deck.js
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const UNIT_OF = { "01_dataset": "Unit_1_2_PBL1_Data_and_Regression", "05_regression": "Unit_1_2_PBL1_Data_and_Regression",
  "02_clustering": "Unit_3_PBL2_Clustering_KMeans", "03_hmm": "Unit_4_PBL2_HMM", "04_comparison": "Unit_4_PBL2_HMM",
  "06_boosting": "Unit_5_PBL3_Gradient_Boosting" };
const IMG = (sub, ...p) => path.join(__dirname, "..", UNIT_OF[sub], "results", sub, ...p);
const OUT = path.join(__dirname, "AI_Supply_Chain_Optimizer_Presentation.pptx");

const THEME = {
  name: "Fresh Grocery",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F2A24", lt1: "FFFFFF", dk2: "1B4332", lt2: "EAF4EE",
    accent1: "2D6A4F", accent2: "E76F51", accent3: "F4A261", accent4: "52B788",
    accent5: "264653", accent6: "8D99AE", hlink: "2D6A4F", folHlink: "6C757D",
  },
};
const HEX = THEME.colors;

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size }));
  const buf = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";                 // 13.333 x 7.5 in
  pres.title = "AI Supply Chain Optimizer";
  pres.author = "Project team";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const W = 13.333, M = 0.6;

  // ---------------------------------------------------------------- layouts
  pres.defineSlideMaster({
    title: "TITLE", background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.0, w: 9.5, h: 1.6, fontSize: 48, bold: true,
          color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "Title" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 3.75, w: 9.5, h: 1.0, fontSize: 20,
          color: C.background2, valign: "top", align: "left", margin: 0 }, text: "Subtitle" } },
    ],
  });
  pres.defineSlideMaster({
    title: "SECTION", background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 2.6, w: 9.0, h: 1.2, fontSize: 40, bold: true,
          color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "Section" } },
      { placeholder: { options: { name: "body", type: "body", x: M, y: 3.95, w: 9.0, h: 1.0, fontSize: 20,
          color: C.background2, valign: "top", align: "left", margin: 0 }, text: "Subtitle" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT", background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.9, fontSize: 30, bold: true,
          color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "Title" } },
    ],
  });

  const content = (title, section) => {
    const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, o));
  const bullets = (s, items, o) => s.addText(items.map((t, i) => ({
      text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })),
    Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top", margin: 0.05 }, o));
  const card = (s, x, y, w, h, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
    { x, y, w, h, rectRadius: 0.12, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: name });
  const img = (s, file, x, y, w, ratio, name) => s.addImage({ path: file, x, y, w, h: w / ratio, objectName: name });
  const stat = (s, x, y, w, value, label, color) => {
    txt(s, value, { x, y, w, h: 0.75, fontSize: 40, bold: true, color: color || C.accent1, fontFace: THEME.headFontFace });
    txt(s, label, { x, y: y + 0.78, w, h: 0.6, fontSize: 13, color: C.text1 });
  };
  const iconCircle = async (s, x, y, d, name, bg) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg }, objectName: "icon circle" });
    s.addImage({ data: await icon(name, "FFFFFF"), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, objectName: name });
  };
  const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    titleFontFace: "+mn-lt", legendFontFace: "+mn-lt", catAxisLabelColor: HEX.dk1, valAxisLabelColor: "6C757D",
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontSize: 12, dataLabelColor: HEX.dk1,
    valGridLine: { color: "E3E8E5", size: 0.75 }, catGridLine: { style: "none" }, titleFontSize: 14, titleColor: HEX.dk1 };

  // ================================================================ 1 title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Introduction" });
  s.addText("AI Supply Chain Optimizer", { placeholder: "title" });
  s.addText("Machine learning for food demand in Indian quick-commerce dark stores", { placeholder: "body" });
  txt(s, "21CSC305P Machine Learning  ·  PBL 1, PBL 2, PBL 3\nTeam: [Name – Reg No]  ·  Faculty: [Name]  ·  November 2026",
    { x: M, y: 5.6, w: 9.5, h: 0.9, fontSize: 14, color: C.background2 });
  await iconCircle(s, 10.6, 2.3, 2.0, "FaStore", C.accent2);
  s.addNotes("Title. Introduce the team and the project: one public quick-commerce dataset used across all five units to plan food stock and riders for every dark store.");

  // ================================================================ 2 problem
  s = content("The problem: plan food stock and riders for every dark store", "Introduction");
  txt(s, "Quick-commerce platforms deliver groceries in about ten minutes from small dark stores. Shelf space is small, fresh food spoils within days and demand changes with neighbourhood, weather and festivals.",
    { x: M, y: 1.4, w: 12.1, h: 0.9, fontSize: 16 });
  const qs = [
    ["FaBoxOpen", C.accent1, "How much?", "Forecast each store's food orders for next week", "PBL 1 (Units 1–2) · PBL 3 (Unit 5)"],
    ["FaMapMarkerAlt", C.accent2, "Where?", "Group stores with the same demand profile", "PBL 2 · Unit 3 (K-Means)"],
    ["FaClock", C.accent5, "When?", "Detect hidden demand regimes over time", "PBL 2 · Unit 4 (HMM)"],
  ];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.13, y = 2.65;
    card(s, x, y, 3.8, 4.1, "question card");
    await iconCircle(s, x + 0.35, y + 0.4, 0.95, qs[i][0], qs[i][1]);
    txt(s, qs[i][2], { x: x + 0.35, y: y + 1.6, w: 3.1, h: 0.6, fontSize: 24, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    txt(s, qs[i][3], { x: x + 0.35, y: y + 2.25, w: 3.1, h: 0.9, fontSize: 16 });
    txt(s, qs[i][4], { x: x + 0.35, y: y + 3.3, w: 3.1, h: 0.5, fontSize: 13, color: C.accent1, bold: true });
  }
  s.addNotes("Overall problem: how can an operator use its order history to plan food replenishment and rider capacity for every dark store? Three linked questions: how much (forecasting), where (store types) and when (demand regimes). All units use the same original dataset; Unit 4 uses a daily sequence derived from it.");

  // ================================================================ 3 dataset
  s = content("One public dataset, cleaned to regular food demand", "Introduction");
  const stats = [["1,95,148", "orders (raw)"], ["58", "dark stores"], ["12", "cities"], ["638", "days, Jan 2025 – Sep 2026"]];
  stats.forEach((v, i) => stat(s, M + (i % 2) * 2.9, 1.55 + Math.floor(i / 2) * 1.75, 2.7, v[0], v[1]));
  txt(s, "Kaggle: India Quick Commerce Orders 2025–2026 (CC0), 39 columns per order", { x: M, y: 5.2, w: 5.6, h: 0.7, fontSize: 13, color: "6C757D" });
  const steps = [["FaDatabase", "Raw orders", "1,95,148"], ["FaFilter", "Keep 7 food categories", "1,17,077"],
                 ["FaGift", "Remove IPL evening effect", "1,13,390"], ["FaTable", "Store-week table, daily sequence", "4,039 / 638 days"]];
  for (let i = 0; i < steps.length; i++) {
    const y = 1.5 + i * 1.3, x = 7.0;
    card(s, x, y, 5.7, 1.0, "pipeline step");
    await iconCircle(s, x + 0.2, y + 0.17, 0.66, steps[i][0], i === 2 ? C.accent2 : C.accent1);
    txt(s, steps[i][1], { x: x + 1.05, y: y + 0.14, w: 2.75, h: 0.72, fontSize: 15, valign: "middle", bold: true });
    txt(s, steps[i][2], { x: x + 3.75, y: y + 0.14, w: 1.8, h: 0.72, fontSize: 15, valign: "middle", align: "right", color: C.accent1, bold: true });
    if (i < steps.length - 1) s.addShape(pres.shapes.DOWN_ARROW, { x: x + 2.7, y: y + 1.03, w: 0.3, h: 0.24, fill: { color: C.accent4 }, line: { color: C.accent4 }, objectName: "arrow" });
  }
  s.addNotes("Dataset and preprocessing. No duplicates or negative amounts; missing values are structural and were kept. Non-food categories removed; IPL excess evening orders removed store by store; result: 1,13,390 clean food orders, a 4,039-row store-week table and a 638-day daily sequence.");

  // ================================================================ 4 IPL
  s = content("IPL matches inflate evening orders, so the excess was removed", "Introduction");
  s.addChart(pres.charts.BAR, [
    { name: "Before removal", labels: ["Regular days", "IPL match days"], values: [30.2, 41.1] },
    { name: "After removal", labels: ["Regular days", "IPL match days"], values: [30.2, 30.4] }],
    Object.assign({ x: M, y: 1.45, w: 6.6, h: 5.4, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.accent6, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"', valAxisMaxVal: 50, valAxisMinVal: 0,
      showLegend: true, legendPos: "b", legendFontSize: 12, showTitle: true, title: "Share of orders placed 7–11 pm (March–May, %)",
      valAxisLabelFormatCode: "0", objectName: "IPL chart" }, chartText));
  bullets(s, [
    "Daytime demand is unchanged on match days (2.59 vs 2.55 orders per store-day)",
    "Evening orders rise from 1.10 to 1.81 per store-day",
    "Per store: normal evening/daytime ratio → expected evening orders → removal probability",
    "3,687 excess orders removed (seed 42); evening share back to the regular level",
  ], { x: 7.7, y: 1.7, w: 5.0, h: 4.6, fontSize: 16 });
  s.addNotes("IPL removal: r_s = evening/daytime on regular March–May days; expected IPL evening = r_s × IPL daytime; p_remove = excess / IPL evening. Each IPL-evening order dropped with that probability.");

  // ================================================================ 5 EDA
  s = content("Demand grows five-fold and peaks late in the evening", "Introduction");
  img(s, IMG("01_dataset", "eda_overview.png"), M, 1.35, 8.3, 1.418, "EDA chart");
  bullets(s, [
    "Daily orders rise from about 60 to 500, with sharp festival spikes — every model must handle growth",
    "The kept food categories include the three largest",
    "Orders peak at 8–10 pm and stay high at night — a rhythm that differs by store",
    "Weekly orders correlate 0.88 with the previous 4-week mean",
  ], { x: 9.2, y: 1.6, w: 3.55, h: 5.2, fontSize: 15 });
  s.addNotes("EDA on the raw data. Growth, category mix and the hourly pattern motivate the features of PBL 1 and the store features of Unit 3.");

  // ================================================================ 6 section PBL1
  pres.addSection({ title: "PBL 1" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "PBL 1" });
  s.addText("PBL 1 · Units 1–2", { placeholder: "title" });
  s.addText("Weekly demand forecasting with Linear, Ridge and Bayesian regression", { placeholder: "body" });
  await iconCircle(s, 10.6, 2.6, 1.7, "FaChartLine", C.accent2);

  // ================================================================ 7 setup
  s = content("Leakage-free features and an honest time split", "PBL 1");
  img(s, path.join(__dirname, "..", "report", "fig_workflow_ch2.png"), M, 1.4, 7.4, 1389 / 760, "workflow");
  card(s, 8.4, 1.4, 4.33, 5.4, "features card");
  txt(s, "Features (known before the week)", { x: 8.65, y: 1.6, w: 3.9, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
  bullets(s, ["lag1, lag2, lag4 and 4-week mean roll4", "store age (growth of new stores)", "festival days in the week",
              "monsoon flag, location type"], { x: 8.65, y: 2.1, w: 3.9, h: 2.0, fontSize: 14 });
  txt(s, "Models", { x: 8.65, y: 4.15, w: 3.9, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
  bullets(s, ["Naive (same as last week)", "Linear regression (OLS)", "Ridge: α by 5-fold time-series CV", "Bayesian ridge: 95% intervals"],
    { x: 8.65, y: 4.6, w: 3.9, h: 2.0, fontSize: 14 });
  txt(s, "Train: 2,531 store-weeks to 23-03-2026  ·  Test: 1,508 store-weeks, 26 unseen weeks  ·  scaler fitted on training only",
    { x: M, y: 5.6, w: 7.4, h: 0.8, fontSize: 14, color: C.accent1, bold: true });
  s.addNotes("Store-week table: 4,039 rows, 58 stores. All features are known at the end of the previous week, so nothing leaks. Split by time, never randomly.");

  // ================================================================ 8 results
  s = content("Regression cuts the forecast error by 21% versus naive", "PBL 1");
  s.addChart(pres.charts.BAR, [{ name: "Test MAE", labels: ["Naive", "Linear", "Ridge", "Bayesian"], values: [6.761, 5.353, 5.353, 5.356] }],
    Object.assign({ x: M, y: 1.45, w: 7.2, h: 5.4, barDir: "col", barGapWidthPct: 55, chartColors: [HEX.accent6, HEX.accent1, HEX.accent1, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 8,
      showLegend: false, showTitle: true, title: "Test MAE (orders per store-week, lower is better)", objectName: "MAE chart" }, chartText));
  stat(s, 8.4, 1.6, 4.3, "0.78", "R² on 26 unseen weeks (naive 0.64)");
  stat(s, 8.4, 3.25, 4.3, "16.4%", "total error as share of demand (WAPE), from 20.7%");
  stat(s, 8.4, 4.9, 4.3, "90.6%", "of test weeks inside the Bayesian 95% interval", C.accent2);
  s.addNotes("Linear, ridge and Bayesian are practically identical: CV chose the smallest ridge penalty (α = 0.001) because data is plentiful relative to ten features. Coverage is below 95% because count noise grows with the level.");

  // ================================================================ 9 forecasts
  s = content("Recent demand dominates; festivals add about 2.6 orders a day", "PBL 1");
  img(s, IMG("05_regression", "forecast_results.png"), M, 1.3, 7.95, 1.359, "forecast chart");
  bullets(s, [
    "roll4 carries most of the signal: +8.4 orders per standard deviation",
    "Each festival day adds about 2.6 orders to a store's week",
    "Naive forecasts are one week late at peaks; regression reacts to the calendar",
    "Slight under-forecast during rapid growth (mean residual +1.67)",
  ], { x: 8.85, y: 1.6, w: 3.9, h: 5.2, fontSize: 15 });
  s.addNotes("Forecast results: scatter of actual vs predicted, Bayesian interval for store DS001, network total per week. Coefficient path shows the OLS solution is already stable.");

  // ================================================================ 10 section PBL2 U3
  pres.addSection({ title: "PBL 2 – Unit 3" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "PBL 2 – Unit 3" });
  s.addText("PBL 2 · Unit 3", { placeholder: "title" });
  s.addText("Clustering the 58 dark stores with K-Means", { placeholder: "body" });
  await iconCircle(s, 10.6, 2.6, 1.7, "FaLayerGroup", C.accent2);

  // ================================================================ 11 features and k
  s = content("Signal-to-noise selects 10 demand features; k = 3", "PBL 2 – Unit 3");
  img(s, IMG("02_clustering", "feature_selection.png"), M, 1.35, 6.2, 1.599, "SNR chart");
  img(s, IMG("02_clustering", "k_selection.png"), 7.1, 1.35, 5.63, 2.638, "k chart");
  bullets(s, [
    "Store rates come from ~2,000 orders, so some differences are pure chance",
    "Kept: volume, night and morning shares, six food categories, basket value",
    "k = 2 only splits off student hubs; k = 3 adds the office stores",
    "K-Means: 50 k-means++ starts, seed 42",
  ], { x: 7.1, y: 3.65, w: 5.63, h: 3.2, fontSize: 15 });
  s.addNotes("SNR = between-store variance / sampling-noise variance. Cancellation, complaints, items and distance are noise. Choice of k uses elbow, silhouette, Davies–Bouldin and Calinski–Harabasz.");

  // ================================================================ 12 clusters
  s = content("Three store types that match location exactly", "PBL 2 – Unit 3");
  img(s, IMG("02_clustering", "clusters_pca_and_profiles.png"), M, 1.4, 7.0, 1.906, "PCA chart");
  txt(s, "PC1 + PC2 explain 78% of the variation; location was not used for clustering", { x: M, y: 5.2, w: 7.0, h: 0.6, fontSize: 13, color: "6C757D" });
  const types = [["FaHome", C.accent1, "Family residential (42)", "18% of orders at 6–10 am; dairy and fresh produce → stock before 6 am"],
                 ["FaMoon", C.accent2, "Late-night snack hubs (11)", "38% of orders at night; snacks in 69% of baskets → night stock and shifts"],
                 ["FaBuilding", C.accent5, "Daytime office (5)", "Fewest night orders; highest basket ₹459 → weekday daytime stock"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.4 + i * 1.85;
    card(s, 7.95, y, 4.78, 1.65, "store type card");
    await iconCircle(s, 8.15, y + 0.4, 0.85, types[i][0], types[i][1]);
    txt(s, types[i][2], { x: 9.2, y: y + 0.15, w: 3.4, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
    txt(s, types[i][3], { x: 9.2, y: y + 0.6, w: 3.4, h: 0.95, fontSize: 13 });
  }
  s.addNotes("Cluster profiles: the clusters reproduce location types exactly — residential, student/PG hubs and office/tech parks — although location was not an input. Service outcomes barely differ between clusters.");

  // ================================================================ 13 evaluation
  s = content("The clusters are well separated and stable", "PBL 2 – Unit 3");
  const ev = [["0.42", "silhouette (no store negative)"], ["0.90", "Davies–Bouldin (lower is better)"],
              ["0.85", "stability: mean ARI over 20 restarts"], ["0.94 / 1.00", "agreement with Ward / GMM (ARI)"]];
  ev.forEach((v, i) => stat(s, M + i * 3.08, 1.45, 2.9, v[0], v[1], i === 2 ? C.accent2 : C.accent1));
  img(s, IMG("02_clustering", "cluster_baskets_and_hours.png"), M + 0.9, 3.05, 10.3, 2.35, "basket chart");
  s.addNotes("Silhouette above 0.25 indicates reasonable structure. Ward agrees almost completely, the Gaussian mixture gives the same partition. Without basket value the silhouette drops to 0.268 and office stores scatter.");

  // ================================================================ 14 section U4
  pres.addSection({ title: "PBL 2 – Unit 4" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "PBL 2 – Unit 4" });
  s.addText("PBL 2 · Unit 4", { placeholder: "title" });
  s.addText("Hidden Markov model of daily demand regimes", { placeholder: "body" });
  await iconCircle(s, 10.6, 2.6, 1.7, "FaProjectDiagram", C.accent2);

  // ================================================================ 15 HMM setup
  s = content("Growth removed: a demand index plus late-delivery rate", "PBL 2 – Unit 4");
  img(s, IMG("01_dataset", "daily_demand_index.png"), M, 1.35, 7.6, 1.722, "demand index chart");
  card(s, 8.55, 1.35, 4.18, 5.5, "HMM card");
  txt(s, "Gaussian HMM  λ = (π, A, B)", { x: 8.8, y: 1.55, w: 3.8, h: 0.5, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, [
    "Observations: demand index (log actual / expected) and late rate",
    "638 consecutive days, one network-level sequence",
    "Baum–Welch (EM), 20 random restarts",
    "BIC chooses 3 hidden states",
    "Viterbi decoding + posterior certainty",
    "Validated on unseen 2026 days",
  ], { x: 8.8, y: 2.1, w: 3.8, h: 4.6, fontSize: 14 });
  s.addNotes("Daily orders grow five-fold, so raw counts would only separate early and late periods. Trend = centred 29-day median, weekday factor from regular days, demand index = log(actual/expected).");

  // ================================================================ 16 regimes
  s = content("Three hidden regimes: normal, monsoon and festival peak", "PBL 2 – Unit 4");
  img(s, IMG("03_hmm", "decoded_regimes.png"), M, 1.35, 7.4, 1.541, "decoded regimes chart");
  const hdr = { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 13, align: "center", valign: "middle" };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 13, color: HEX.dk1, align: "center", valign: "middle" }, o || {}) });
  s.addTable([
    [{ text: "Regime", options: hdr }, { text: "Days", options: hdr }, { text: "Late", options: hdr }, { text: "Stay", options: hdr }],
    [cell("Normal", { align: "left" }), cell("357"), cell("22.5%"), cell("35 d")],
    [cell("Monsoon", { align: "left" }), cell("258"), cell("32.3%"), cell("54 d")],
    [cell("Festival peak", { align: "left" }), cell("23"), cell("29.6%"), cell("2.4 d")],
  ], { x: 8.3, y: 1.45, w: 4.43, colW: [1.55, 0.85, 1.0, 1.03], rowH: 0.42, border: { type: "solid", pt: 0.75, color: "CFD8D3" }, objectName: "regime table" });
  bullets(s, [
    "Found without weather or festival labels",
    "Normal Nov–May; Monsoon Jun–Oct (half of orders in rain)",
    "Festival peaks: +32% demand, last ~2 days",
    "Tomorrow (1 Oct 2026): monsoon 0.978, festival 0.018",
  ], { x: 8.3, y: 3.45, w: 4.43, h: 3.4, fontSize: 14 });
  s.addNotes("Expected stay = 1/(1 − a_ii). Normal and Monsoon have the same demand distribution but differ in lateness; Festival peak returns to Normal with probability 0.279.");

  // ================================================================ 17 validation
  s = content("The HMM gives stable regimes and predicts unseen days best", "PBL 2 – Unit 4");
  s.addChart(pres.charts.BAR, [{ name: "Regime switches", labels: ["K-Means", "GMM", "HMM"], values: [251, 270, 19] }],
    Object.assign({ x: M, y: 1.45, w: 5.0, h: 4.6, barDir: "col", barGapWidthPct: 55, chartColors: [HEX.accent6, HEX.accent6, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", valAxisMinVal: 0, valAxisMaxVal: 320,
      showLegend: false, showTitle: true, title: "Label switches over 638 days", objectName: "switch chart" }, chartText));
  s.addChart(pres.charts.BAR, [{ name: "Log-likelihood", labels: ["Single Gaussian", "GMM (no transitions)", "HMM"], values: [2.017, 2.182, 2.501] }],
    Object.assign({ x: 5.85, y: 1.45, w: 4.3, h: 4.6, barDir: "col", barGapWidthPct: 55, chartColors: [HEX.accent6, HEX.accent6, HEX.accent1],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 3,
      showLegend: false, showTitle: true, title: "Log-likelihood per unseen 2026 day", objectName: "oos chart" }, chartText));
  stat(s, 10.45, 1.6, 2.3, "82%", "festival days decoded as peaks (18 of 22)", C.accent2);
  stat(s, 10.45, 3.25, 2.3, "94.7%", "of Jun–Sep days in the monsoon regime");
  stat(s, 10.45, 4.9, 2.3, "97.7%", "mean posterior certainty");
  txt(s, "Same daily observations for all methods; HMM trained on 2025 and scored on 2026", { x: M, y: 6.25, w: 9.5, h: 0.5, fontSize: 13, color: "6C757D" });
  s.addNotes("Time-blind methods flip labels on single rainy-looking days. Only 5 regular days are false festival peaks (noisy first weeks of 2025).");

  // ================================================================ 18 section PBL3
  pres.addSection({ title: "PBL 3" });
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "PBL 3" });
  s.addText("PBL 3 · Unit 5", { placeholder: "title" });
  s.addText("Ensemble learning: gradient boosting on the PBL 1 task", { placeholder: "body" });
  await iconCircle(s, 10.6, 2.6, 1.7, "FaTree", C.accent2);

  // ================================================================ 19 boosting
  s = content("Trees can't extrapolate growth: predict the relative change", "PBL 3");
  img(s, IMG("06_boosting", "boosting_feature_analysis.png"), M, 1.35, 7.6, 2.35, "boosting analysis chart");
  stat(s, M, 4.75, 3.4, "19.2%", "of test store-weeks exceed the store's training maximum", C.accent2);
  card(s, 4.4, 4.65, 3.8, 2.15, "target card");
  txt(s, "Growth-free target", { x: 4.6, y: 4.8, w: 3.4, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
  txt(s, "z = log(orders / roll4)\nŷ = roll4 · exp(F(x))", { x: 4.6, y: 5.3, w: 3.4, h: 1.3, fontSize: 16, color: C.text1 });
  card(s, 8.55, 1.35, 4.18, 5.45, "settings card");
  txt(s, "Configuration", { x: 8.8, y: 1.55, w: 3.8, h: 0.5, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, [
    "GradientBoostingRegressor, CART trees, subsample 0.8",
    "Grid: 200/400/800 trees × rate 0.02/0.05 × depth 2/3",
    "5-fold time-series CV on training weeks only",
    "Best: 200 trees, depth 2, rate 0.02",
    "Same split, test weeks and metrics as PBL 1",
  ], { x: 8.8, y: 2.1, w: 3.8, h: 4.5, fontSize: 14 });
  s.addNotes("A tree only predicts values inside its training range. Dividing out the recent level makes the target stationary: mean +0.011, std 0.270.");

  // ================================================================ 20 comparison
  s = content("Growth-free boosting is the most accurate model", "PBL 3");
  s.addChart(pres.charts.BAR, [{ name: "Test MAE",
      labels: ["Naive", "GB (original features)", "Bayesian", "Ridge", "Linear", "GB (growth-free)"],
      values: [6.761, 5.917, 5.356, 5.353, 5.353, 5.315] }],
    Object.assign({ x: M, y: 1.45, w: 7.6, h: 5.4, barDir: "bar", barGapWidthPct: 45,
      chartColors: [HEX.accent6, HEX.accent6, HEX.accent1, HEX.accent1, HEX.accent1, HEX.accent2],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 8,
      showLegend: false, showTitle: true, title: "Test MAE on the same 1,508 store-weeks (lower is better)", objectName: "comparison chart" }, chartText));
  stat(s, 8.6, 1.6, 4.1, "5.32", "MAE of growth-free boosting (R² 0.784), best on every metric", C.accent2);
  stat(s, 8.6, 3.25, 4.1, "4.34", "Poisson noise floor: the best possible MAE");
  stat(s, 8.6, 4.9, 4.1, "0.7%", "gain over linear regression — the signal is mostly linear");
  s.addNotes("Plain boosting on order levels is second worst because it cannot predict weeks above the training range. Learning curve: validation MAE is lowest at 133 trees and rises by 800 — the CV choice of 200 trees avoids over-fitting. Recommendation: Bayesian regression for weekly replenishment (intervals), boosting once richer features arrive.");

  // ================================================================ 21 integrated
  pres.addSection({ title: "Conclusion" });
  s = content("Monsoon, not store type, drives late deliveries", "Conclusion");
  const late = [["Family residential", 23.1, 32.8, 31.3], ["Late-night snack hubs", 22.6, 31.6, 29.1], ["Daytime office", 22.1, 33.4, 24.4]];
  const shade = v => (v >= 31 ? HEX.accent2 : v >= 27 ? HEX.accent3 : HEX.lt2);
  const rows = [[{ text: "Late deliveries (%)", options: hdr }, { text: "Normal", options: hdr }, { text: "Monsoon", options: hdr }, { text: "Festival", options: hdr }]];
  late.forEach(r => rows.push([cell(r[0], { align: "left", bold: true })].concat(r.slice(1).map(v =>
    cell(v.toFixed(1), { fill: { color: shade(v) }, bold: v >= 31, color: v >= 31 ? HEX.lt1 : HEX.dk1, fontSize: 16 })))));
  s.addTable(rows, { x: M, y: 1.6, w: 6.9, colW: [2.7, 1.4, 1.4, 1.4], rowH: 0.75, border: { type: "solid", pt: 1, color: "FFFFFF" }, objectName: "late table" });
  txt(s, "Every order labelled with its store's cluster (Unit 3) and its day's regime (Unit 4)", { x: M, y: 4.85, w: 6.9, h: 0.6, fontSize: 13, color: "6C757D" });
  bullets(s, [
    "Monsoon adds 9–11 points of late deliveries in every store type → rider buffers Jun–Oct",
    "Festival peaks hit residential and student stores (+6–8 points) far more than offices (+2)",
    "Store types differ in what and when people buy, not in service quality",
    "Forecasts sit near the noise floor: new data (prices, promotions, weather) matters more than model complexity",
  ], { x: 7.9, y: 1.6, w: 4.83, h: 5.2, fontSize: 15 });
  s.addNotes("Units 3 and 4 use different data units (stores vs days), so they are combined rather than compared. Units 1, 2 and 5 share task, test weeks and metrics and are compared directly.");

  // ================================================================ 22 conclusion
  s = pres.addSlide({ masterName: "SECTION", sectionTitle: "Conclusion" });
  s.addText("Conclusion and future scope", { placeholder: "title" });
  s.addText("How much to send, to which store type, under which conditions", { placeholder: "body" });
  const concl = [["Forecasting", "MAE 6.76 → 5.35 → 5.32; Bayesian intervals for safety stock"],
                 ["Store types", "3 stable types (silhouette 0.42) with their own plans"],
                 ["Regimes", "Normal, monsoon, festival found without labels; 82% recall"],
                 ["Next steps", "Real SKU data, prices and weather; regime-aware forecasts; weekly retraining"]];
  concl.forEach((c, i) => {
    const x = M + i * 3.08;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.15, w: 2.9, h: 1.75, rectRadius: 0.12, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "conclusion card" });
    txt(s, c[0], { x: x + 0.2, y: 5.28, w: 2.5, h: 0.4, fontSize: 16, bold: true, color: i === 3 ? C.accent3 : C.background1 });
    txt(s, c[1], { x: x + 0.2, y: 5.7, w: 2.5, h: 1.1, fontSize: 13, color: C.background2 });
  });
  s.addNotes("Summary of achievements and future scope: real product-level data, HMM regime probabilities and weather as forecast features, per-city HMMs, and a weekly retraining pipeline with alerts for store managers.");

  // ================================================================ 23 references
  s = content("Dataset and references", "Conclusion");
  await iconCircle(s, M, 1.5, 0.8, "FaDatabase", C.accent1);
  txt(s, [{ text: "Dataset: ", options: { bold: true } },
          { text: "India Quick Commerce Orders 2025–2026, Kaggle (CC0) — kaggle.com/datasets/abirambharathisa/india-quick-commerce-orders-2025-2026" }],
    { x: M + 1.0, y: 1.5, w: 11.1, h: 0.8, fontSize: 15, valign: "middle" });
  bullets(s, [
    "F. Pedregosa et al., \"Scikit-learn: Machine learning in Python,\" JMLR, 2011; hmmlearn documentation",
    "T. Hastie, R. Tibshirani, J. Friedman, The Elements of Statistical Learning, 2nd ed., Springer, 2009",
    "C. M. Bishop, Pattern Recognition and Machine Learning, Springer, 2006",
    "A. E. Hoerl, R. W. Kennard, \"Ridge regression,\" Technometrics, 1970",
    "S. P. Lloyd, \"Least squares quantization in PCM,\" IEEE Trans. Inf. Theory, 1982; P. J. Rousseeuw, \"Silhouettes,\" 1987",
    "L. R. Rabiner, \"A tutorial on hidden Markov models,\" Proc. IEEE, 1989; G. Schwarz, \"Estimating the dimension of a model,\" 1978",
    "J. H. Friedman, \"Greedy function approximation: A gradient boosting machine,\" Ann. Stat., 2001",
    "Code: six Jupyter notebooks (01–06) in the project folder; Colab links: [add links]",
  ], { x: M, y: 2.6, w: 12.1, h: 4.3, fontSize: 15 });
  s.addNotes("Sources used in the project and the report.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

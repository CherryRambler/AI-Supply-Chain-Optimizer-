// Gradient boosting deck (same design as the project deck). Run: node build_deck_gb.js
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
const OUT = path.join(__dirname, "AI_Supply_Chain_Optimizer_Gradient_Boosting.pptx");

const THEME = {
  name: "Fresh Grocery", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1F2A24", lt1: "FFFFFF", dk2: "1B4332", lt2: "EAF4EE", accent1: "2D6A4F", accent2: "E76F51",
            accent3: "F4A261", accent4: "52B788", accent5: "264653", accent6: "8D99AE", hlink: "2D6A4F", folHlink: "6C757D" },
};
const HEX = THEME.colors;

async function icon(name, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(fa[name], { color: "#" + color, size }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer()).toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title = "AI Supply Chain Optimizer – Gradient Boosting";
  pres.author = "Project team";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const W = 13.333, M = 0.6;

  pres.defineSlideMaster({ title: "TITLE", background: { color: C.text2 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 1.9, w: 9.6, h: 1.7, fontSize: 48, bold: true,
        color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "Title" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 3.75, w: 9.6, h: 1.1, fontSize: 20,
        color: C.background2, valign: "top", align: "left", margin: 0 }, text: "Subtitle" } }] });
  pres.defineSlideMaster({ title: "DARK", background: { color: C.text2 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.9, fontSize: 30, bold: true,
        color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "Title" } }] });
  pres.defineSlideMaster({ title: "CONTENT", background: { color: C.background1 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.35, w: W - 2 * M, h: 0.9, fontSize: 30, bold: true,
        color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "Title" } }] });

  const content = (title, section, master = "CONTENT") => {
    const s = pres.addSlide({ masterName: master, sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    return s;
  };
  const txt = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top", margin: 0 }, o));
  const bullets = (s, items, o) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })),
    Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top", margin: 0.05 }, o));
  const card = (s, x, y, w, h, name, color) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
    { x, y, w, h, rectRadius: 0.12, fill: { color: color || C.background2 }, line: { color: color || C.background2 }, objectName: name });
  const img = (s, file, x, y, w, ratio, name) => s.addImage({ path: file, x, y, w, h: w / ratio, objectName: name });
  const stat = (s, x, y, w, value, label, color) => {
    txt(s, value, { x, y, w, h: 0.75, fontSize: 40, bold: true, color: color || C.accent1, fontFace: THEME.headFontFace });
    txt(s, label, { x, y: y + 0.78, w, h: 0.6, fontSize: 13 });
  };
  const iconCircle = async (s, x, y, d, name, bg) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg }, objectName: "icon circle" });
    s.addImage({ data: await icon(name, "FFFFFF"), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, objectName: name });
  };
  const arrowR = (s, x, y) => s.addShape(pres.shapes.RIGHT_ARROW, { x, y, w: 0.32, h: 0.3, fill: { color: C.accent4 }, line: { color: C.accent4 }, objectName: "arrow" });
  const chartText = { catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
    titleFontFace: "+mn-lt", legendFontFace: "+mn-lt", catAxisLabelColor: HEX.dk1, valAxisLabelColor: "6C757D",
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontSize: 12, dataLabelColor: HEX.dk1,
    valGridLine: { color: "E3E8E5", size: 0.75 }, catGridLine: { style: "none" }, titleFontSize: 14, titleColor: HEX.dk1 };
  const hdr = { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 14, align: "center", valign: "middle" };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 14, color: HEX.dk1, align: "center", valign: "middle" }, o || {}) });


  // ================================================================ 1 title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Introduction" });
  s.addText("AI Supply Chain Optimizer", { placeholder: "title" });
  s.addText("Forecasting weekly food demand of quick-commerce dark stores with gradient boosting", { placeholder: "body" });
  txt(s, "How gradient boosting works, how we adapted it to fast-growing demand, and how it compares with regression",
    { x: M, y: 4.85, w: 9.6, h: 0.7, fontSize: 16, italic: true, color: C.accent3 });
  txt(s, "21CSC305P Machine Learning  ·  PBL 3\nTeam: [Name – Reg No]  ·  Faculty: [Name]  ·  November 2026",
    { x: M, y: 5.8, w: 9.6, h: 0.9, fontSize: 14, color: C.background2 });
  await iconCircle(s, 10.6, 2.3, 2.0, "FaTree", C.accent2);
  s.addNotes("This part of the project re-solves the weekly demand forecasting task with an ensemble method, gradient boosting, and compares it with the regression models on exactly the same data, test weeks and metrics.");

  // ================================================================ 2 task
  s = content("The task: forecast next week's food orders for every dark store", "Introduction");
  img(s, path.join(__dirname, "..", "report", "fig_workflow_unit5.png"), M, 1.45, 7.6, 1389 / 760, "workflow");
  const st2 = (y, v, l, col) => {
    txt(s, v, { x: 8.75, y, w: 4.0, h: 0.75, fontSize: 40, bold: true, color: col || C.accent1, fontFace: THEME.headFontFace });
    txt(s, l, { x: 8.75, y: y + 0.78, w: 4.0, h: 0.75, fontSize: 14 });
  };
  st2(1.45, "4,039", "store-weeks: 58 dark stores, each followed week by week (Kaggle, India Quick Commerce Orders 2025–2026)");
  st2(3.15, "26", "test weeks (30-03-2026 to 21-09-2026) kept hidden during training, so the model is judged on the future");
  st2(4.85, "5.35", "MAE of linear regression: on average its forecast is off by 5.35 orders. This is the score to beat", C.accent2);
  txt(s, [
    { text: "What the model sees: ", options: { bold: true, color: HEX.accent1 } },
    { text: "only facts known before the week starts — orders 1, 2 and 4 weeks ago, the 4-week average, store age, festival days, monsoon and location type. Same features, split and metrics as the regression models, so the comparison is fair." }],
    { x: M, y: 5.8, w: 7.6, h: 1.25, fontSize: 14 });
  s.addNotes("Each row is one store in one week. A dark store is a small delivery-only warehouse. MAE (mean absolute error) is the average size of the forecast error, measured in orders.");

  // ================================================================ 3 why
  s = content("Why gradient boosting suits this problem", "Introduction");
  const why = [["FaProjectDiagram", C.accent1, "Non-linear effects", "Demand does not change by a fixed amount. A festival in October lifts orders more than one in March. Trees learn such combined effects on their own."],
               ["FaTable", C.accent5, "Built for tabular data", "Our data is a table of numbers and yes/no columns. Trees use it directly: no scaling and no need to write the formula by hand."],
               ["FaChartLine", C.accent2, "Controllable fit", "Three simple settings — number of trees, learning rate and tree depth — let us decide how much detail the model is allowed to learn."]];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.13, y = 1.5;
    card(s, x, y, 3.8, 4.1, "reason card");
    await iconCircle(s, x + 0.35, y + 0.35, 0.9, why[i][0], why[i][1]);
    txt(s, why[i][2], { x: x + 0.35, y: y + 1.45, w: 3.2, h: 0.55, fontSize: 21, bold: true, color: C.text2, fontFace: THEME.headFontFace });
    txt(s, why[i][3], { x: x + 0.35, y: y + 2.05, w: 3.15, h: 1.9, fontSize: 16 });
  }
  txt(s, [
    { text: "In one line: ", options: { bold: true } },
    { text: "linear regression adds a fixed amount for each feature; gradient boosting builds many small decision trees, each one fixing the mistakes of the trees before it.", options: { breakLine: true } },
    { text: "Our question: does this flexible ensemble beat simple linear regression on the same forecasting task?", options: { italic: true, color: HEX.accent1 } }],
    { x: M, y: 5.8, w: 12.1, h: 1.15, fontSize: 15 });
  s.addNotes("Linear regression assumes each feature adds a fixed amount. Gradient boosting can learn, for example, that festivals matter more in the autumn festival season.");

  // ================================================================ 4 how it works
  pres.addSection({ title: "How gradient boosting works" });
  s = content("How gradient boosting works, step by step", "How gradient boosting works");
  const steps = [["1", "Start simple", "Give every store-week the same first guess: the average. This guess is called F₀."],
                 ["2", "Measure errors", "Residual = actual − prediction. A positive residual means we predicted too low."],
                 ["3", "Fit a small tree", "A shallow decision tree learns to predict the residuals: where we are wrong, and by how much."],
                 ["4", "Take a small step", "Add the tree's output × learning rate ν (0.02) to the prediction, fixing only a little."],
                 ["5", "Repeat M times", "Each new tree fixes what is still wrong. Many small fixes add up to an accurate forecast."]];
  for (let i = 0; i < 5; i++) {
    const x = M + i * 2.48, y = 1.45;
    card(s, x, y, 2.16, 3.1, "step card");
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fill: { color: i === 4 ? C.accent2 : C.accent1 }, line: { color: i === 4 ? C.accent2 : C.accent1 }, objectName: "step number" });
    txt(s, steps[i][0], { x: x + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle" });
    txt(s, steps[i][1], { x: x + 0.2, y: y + 0.88, w: 1.85, h: 0.4, fontSize: 16, bold: true, color: C.text2 });
    txt(s, steps[i][2], { x: x + 0.2, y: y + 1.32, w: 1.82, h: 1.7, fontSize: 14 });
    if (i < 4) arrowR(s, x + 2.17, y + 1.4);
  }
  card(s, M, 4.75, 12.13, 2.3, "formula card", C.text2);
  txt(s, "Final model = first guess + many small corrections", { x: M + 0.35, y: 4.9, w: 7.5, h: 0.45, fontSize: 16, bold: true, color: C.accent3 });
  const fr = (t, sub) => ({ text: t, options: sub ? { subscript: true } : {} });
  txt(s, [fr("F(x) = F"), fr("0", 1), fr(" + ν·h"), fr("1", 1), fr("(x) + ν·h"), fr("2", 1), fr("(x) + … + ν·h"), fr("M", 1), fr("(x)")],
    { x: M + 0.35, y: 5.35, w: 7.4, h: 0.65, fontSize: 26, bold: true, color: C.background1, fontFace: THEME.headFontFace });
  txt(s, "hₘ = the m-th tree, trained on the residuals rₘ = y − Fₘ₋₁(x);  ν = learning rate.\nWhy \"gradient\"? For squared error the residual is exactly the negative gradient of the loss, so each tree is one small gradient-descent step.",
    { x: M + 0.35, y: 6.05, w: 11.4, h: 0.9, fontSize: 14, color: C.background2 });
  txt(s, "Our model: M = 200 trees, depth 2, ν = 0.02", { x: 8.3, y: 5.4, w: 4.2, h: 0.55, fontSize: 16, bold: true, color: C.accent3, align: "right" });
  s.addNotes("Gradient boosting is gradient descent in function space: with squared-error loss the negative gradient is simply the residual, so each tree is trained to predict the current errors. The learning rate shrinks each tree's contribution so that many small corrections add up smoothly. Each tree is trained on a random 80% of the rows (stochastic boosting).");

  // ================================================================ 5 one tree
  s = content("Each tree is a small set of if–then rules", "How gradient boosting works");
  txt(s, "The first tree of our trained model (depth 2, learned from 2,024 sampled store-weeks)", { x: M, y: 1.3, w: 8.4, h: 0.4, fontSize: 14, color: "6C757D" });
  const node = (x, y, w, t) => { card(s, x, y, w, 0.75, "split node", C.accent1); txt(s, t, { x, y, w, h: 0.75, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle" }); };
  const leaf = (x, y, v, n, hot) => { card(s, x, y, 1.85, 1.0, "leaf", hot ? C.accent2 : C.background2);
    txt(s, [{ text: v, options: { bold: true, fontSize: 18, breakLine: true } }, { text: n, options: { fontSize: 12 } }],
      { x, y: y + 0.05, w: 1.85, h: 0.9, align: "center", valign: "middle", color: hot ? C.background1 : C.text1 }); };
  const line = (x1, y1, x2, y2) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: y1, w: Math.abs(x2 - x1) || 0.01, h: y2 - y1, flipH: x2 < x1, line: { color: "8D99AE", width: 1.5 }, objectName: "edge" });
  node(3.0, 1.85, 3.4, "festival_days ≤ 1.5 ?");
  node(0.9, 3.15, 2.9, "r4 ≤ 1.78 ?");
  node(5.6, 3.15, 2.9, "week of year ≤ 37 ?");
  line(4.0, 2.6, 2.35, 3.15); line(5.4, 2.6, 7.05, 3.15);
  txt(s, "yes", { x: 2.6, y: 2.68, w: 0.6, h: 0.3, fontSize: 12, color: "6C757D" }); txt(s, "no", { x: 6.4, y: 2.68, w: 0.6, h: 0.3, fontSize: 12, color: "6C757D" });
  leaf(0.6, 4.5, "−0.023", "1,770 weeks"); leaf(2.55, 4.5, "−1.266", "2 weeks");
  leaf(5.3, 4.5, "+0.153", "214 weeks", true); leaf(7.25, 4.5, "+0.385", "38 weeks", true);
  line(1.8, 3.9, 1.5, 4.5); line(2.9, 3.9, 3.45, 4.5); line(6.5, 3.9, 6.2, 4.5); line(7.6, 3.9, 8.15, 4.5);
  txt(s, [
    { text: "How to read it: ", options: { bold: true, color: HEX.accent1 } },
    { text: "start at the top and answer each question — go left for yes, right for no. The box you end in (a leaf) gives this tree's correction for that store-week. Orange leaves raise the forecast; grey leaves leave it almost unchanged." }],
    { x: M, y: 5.75, w: 8.5, h: 1.2, fontSize: 15 });
  card(s, 9.45, 1.4, 3.28, 5.6, "meaning card");
  txt(s, "What it learned", { x: 9.7, y: 1.6, w: 2.9, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, [
    "First question: does the week have 2 or more festival days? If yes, demand needs an upward correction",
    "Even more after week 37 — the Dussehra–Diwali season (+0.385)",
    "Leaf values are corrections to log(orders ÷ roll4); +0.153 means roughly 16% above the recent average",
    "Each value is multiplied by ν = 0.02, so one tree moves a forecast only slightly",
  ], { x: 9.7, y: 2.1, w: 2.9, h: 4.8, fontSize: 15 });
  s.addNotes("Real splits from estimators_[0] of the saved model. r4 = orders four weeks ago divided by the 4-week mean. The −1.266 leaf covers only two unusual weeks. Later trees pick up other patterns, e.g. small stores (roll4 ≤ 8).");

  // ================================================================ 6 worked example
  s = content("Worked example: one festival week, forecast tree by tree", "How gradient boosting works");
  const pts = [22.29, 23.19, 23.95, 24.55, 25.02, 25.41, 25.53, 25.95, 26.09, 26.24, 26.33, 26.38, 26.44, 26.52, 26.75, 26.94, 26.93, 26.99, 27.09, 27.13, 27.22];
  const labs = pts.map((_, i) => String(i * 10));
  s.addChart(pres.charts.LINE, [
    { name: "Boosting forecast", labels: labs, values: pts },
    { name: "Actual orders (28)", labels: labs, values: pts.map(() => 28) }],
    Object.assign({ x: M, y: 1.45, w: 7.3, h: 5.5, chartColors: [HEX.accent1, HEX.accent2], lineSize: 3, lineDataSymbol: "none",
      valAxisMinVal: 20, valAxisMaxVal: 30, showLegend: true, legendPos: "b", legendFontSize: 12, showTitle: true,
      title: "Forecast for store DS052, week of 24-08-2026, after m trees", catAxisTitle: "number of trees", showCatAxisTitle: true,
      catAxisTitleFontSize: 12, catAxisTitleColor: "6C757D", objectName: "staged chart" }, chartText));
  const ex = [["22.3", "Start: the store's 4-week average (22) × the first guess e^F₀ — before any tree"],
              ["25.4", "After 50 trees: each festival-aware tree has added a small upward correction"],
              ["27.2", "After all 200 trees — very close to the actual 28 orders (last week's value, 21, was far off)"]];
  ex.forEach((e, i) => {
    const y = 1.45 + i * 1.6;
    txt(s, e[0], { x: 8.4, y, w: 4.3, h: 0.7, fontSize: 36, bold: true, color: i === 2 ? C.accent2 : C.accent1, fontFace: THEME.headFontFace });
    txt(s, e[1], { x: 8.4, y: y + 0.72, w: 4.3, h: 0.8, fontSize: 14 });
  });
  txt(s, "This week had 3 festival days. The first tree adds only 0.02 × 0.153 ≈ 0.003; 200 such small steps lift the forecast by 22%.",
    { x: 8.4, y: 6.25, w: 4.3, h: 0.8, fontSize: 14, italic: true, color: C.accent1 });
  s.addNotes("Computed with staged_predict on the saved model. The first tree adds 0.02 × 0.153 = 0.003 to the log ratio; 200 small steps move the forecast from 22.3 to 27.2 orders. The naive forecast (last week) was 21.");

  // ================================================================ 7 settings
  s = content("Four settings control how boosting learns", "How gradient boosting works");
  const hp = [["FaLayerGroup", "Number of trees", "How many correction steps the model takes. Too few and it misses patterns; too many and it starts to memorise random noise.", "Tried 200, 400, 800 → chose 200"],
              ["FaChartLine", "Learning rate ν", "How much of each tree's correction is kept. A small rate means careful steps: smoother and safer, but more trees are needed.", "Tried 0.02, 0.05 → chose 0.02"],
              ["FaTree", "Tree depth", "How many questions each tree may ask. Depth 2 = two questions, so a tree can combine two features, e.g. festival and season.", "Tried 2, 3 → chose 2"],
              ["FaFilter", "Subsample", "Each tree is trained on a random 80% of the rows. The trees differ a little, so the model relies less on any single week.", "Fixed at 0.8"]];
  for (let i = 0; i < 4; i++) {
    const x = M + (i % 2) * 6.15, y = 1.45 + Math.floor(i / 2) * 2.8;
    card(s, x, y, 5.95, 2.6, "setting card");
    await iconCircle(s, x + 0.3, y + 0.3, 0.85, hp[i][0], i === 1 ? C.accent2 : C.accent1);
    txt(s, hp[i][1], { x: x + 1.4, y: y + 0.25, w: 4.3, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    txt(s, hp[i][2], { x: x + 1.4, y: y + 0.8, w: 4.3, h: 1.2, fontSize: 15 });
    txt(s, hp[i][3], { x: x + 1.4, y: y + 2.05, w: 4.3, h: 0.4, fontSize: 15, bold: true, color: C.accent1 });
  }
  s.addNotes("These four values define the model. They were chosen by cross-validation on the training weeks only (slide 10).");

  // ================================================================ 8 problem
  pres.addSection({ title: "Our model" });
  s = content("Problem: trees cannot predict beyond their training range", "Our model");
  img(s, IMG("06_boosting", "boosting_feature_analysis.png"), M, 1.4, 7.9, 2.35, "extrapolation chart");
  stat(s, 9.0, 1.45, 3.7, "19.2%", "of test store-weeks are higher than the store's best training week", C.accent2);
  stat(s, 9.0, 3.0, 3.7, "5.92", "MAE of plain boosting on order counts — worse than linear (5.35)");
  card(s, 9.0, 4.55, 3.73, 2.45, "definition card");
  txt(s, [
    { text: "Extrapolation", options: { bold: true, fontSize: 16, color: HEX.dk2, breakLine: true } },
    { text: "means predicting outside the range seen in training. A straight line can be extended upwards; a tree can only repeat values it has seen." }],
    { x: 9.25, y: 4.7, w: 3.3, h: 2.2, fontSize: 15 });
  bullets(s, [
    "A leaf predicts the average of the training weeks that land in it, so a tree can never forecast more orders than it has already seen",
    "Quick-commerce demand keeps growing, so about one in five test weeks is a new record for its store",
    "Result: plain boosting under-predicts these weeks and loses to simple linear regression",
  ], { x: M, y: 4.95, w: 7.9, h: 2.05, fontSize: 14 });
  s.addNotes("Left chart: each dot is a store; dots near or above the diagonal have test weeks close to or above their training maximum. This is a known limitation of all tree models. Right chart: the growth-free target used on the next slide.");

  // ================================================================ 9 fix
  s = content("Fix: predict the relative change, not the level", "Our model");
  const flow = [["1  Recent level", "roll4 = average orders of the last 4 weeks. It already carries the store's size and growth.", C.accent5],
                ["2  Boosting target", "z = log(orders ÷ roll4): how far this week is from the recent level (0 = same, +0.1 ≈ 10% more).", C.accent1],
                ["3  Back to orders", "ŷ = roll4 × e^F(x): the recent level times the predicted change gives the order forecast.", C.accent2]];
  for (let i = 0; i < 3; i++) {
    const x = M + i * 4.13, y = 1.45;
    card(s, x, y, 3.8, 2.3, "flow card", flow[i][2]);
    txt(s, flow[i][0], { x: x + 0.3, y: y + 0.2, w: 3.2, h: 0.5, fontSize: 19, bold: true, color: C.background1 });
    txt(s, flow[i][1], { x: x + 0.3, y: y + 0.78, w: 3.25, h: 1.4, fontSize: 15, color: C.background1 });
    if (i < 2) arrowR(s, x + 3.84, y + 1.0);
  }
  card(s, M, 4.0, 5.95, 3.0, "features card");
  txt(s, "Relative features added", { x: M + 0.3, y: 4.15, w: 5.4, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, ["r1, r2, r4 = orders 1, 2 and 4 weeks ago ÷ roll4: recent momentum on the same scale for big and small stores",
              "woy = week of the year, so trees can learn the season",
              "Kept: roll4, store age, festival days, monsoon, location"], { x: M + 0.3, y: 4.65, w: 5.4, h: 2.25, fontSize: 15 });
  card(s, 6.75, 4.0, 5.98, 3.0, "result card");
  txt(s, "Why it works", { x: 7.05, y: 4.15, w: 5.4, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, ["Growth is divided out, so test weeks look like training weeks: nothing new to extrapolate",
              "The target is stable: mean +0.011, spread 0.270, nearly symmetric",
              "Trees learn only the deviation; the level comes from roll4, which grows with the store"], { x: 7.05, y: 4.65, w: 5.4, h: 2.25, fontSize: 15 });
  s.addNotes("This reformulation is the key design decision of this part of the project: it turns a growing series into a stationary one that trees can model.");

  // ================================================================ 10 tuning
  s = content("Tuning with time-series cross-validation", "Our model");
  txt(s, "5 folds, each trained on earlier weeks and validated on the next block", { x: M, y: 1.4, w: 6.0, h: 0.45, fontSize: 15, bold: true, color: C.text2 });
  for (let f = 0; f < 5; f++) {
    const y = 1.95 + f * 0.5, tw = 1.6 + f * 0.75;
    s.addShape(pres.shapes.RECTANGLE, { x: M + 0.6, y, w: tw, h: 0.36, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "train block" });
    s.addShape(pres.shapes.RECTANGLE, { x: M + 0.6 + tw + 0.05, y, w: 0.7, h: 0.36, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "validation block" });
    txt(s, "Fold " + (f + 1), { x: M, y, w: 0.6, h: 0.36, fontSize: 12, valign: "middle", color: "6C757D" });
  }
  txt(s, [{ text: "■ ", options: { color: HEX.accent1 } }, { text: "train   " }, { text: "■ ", options: { color: HEX.accent2 } }, { text: "validate   → time" }],
    { x: M + 0.6, y: 4.5, w: 5.0, h: 0.4, fontSize: 13, color: "6C757D" });
  txt(s, [
    { text: "What is cross-validation? ", options: { bold: true, color: HEX.accent1 } },
    { text: "Each candidate setting is trained on part of the training weeks and scored on weeks it has not seen. In the time-series version the scored weeks always come after the training weeks, just like real forecasting, so no future information leaks into the choice." }],
    { x: M, y: 5.05, w: 6.0, h: 1.95, fontSize: 14 });
  const cv = [[200, 0.02, 2, 0.2167], [200, 0.02, 3, 0.2189], [400, 0.02, 2, 0.2190], [200, 0.05, 2, 0.2212], [400, 0.02, 3, 0.2227], [800, 0.02, 2, 0.2233], [800, 0.05, 3, 0.2409]];
  const rows = [[{ text: "Trees", options: hdr }, { text: "Rate ν", options: hdr }, { text: "Depth", options: hdr }, { text: "CV MAE", options: hdr }]];
  cv.forEach((r, i) => {
    const best = i === 0, last = i === cv.length - 1;
    const o = best ? { bold: true, fill: { color: HEX.lt2 }, color: HEX.accent1 } : {};
    rows.push([cell(last ? "… " + r[0] : String(r[0]), o), cell(String(r[1]), o), cell(String(r[2]), o), cell(r[3].toFixed(4), o)]);
  });
  s.addTable(rows, { x: 7.1, y: 1.4, w: 5.63, colW: [1.4, 1.4, 1.3, 1.53], rowH: 0.44, border: { type: "solid", pt: 0.75, color: "CFD8D3" }, objectName: "cv table" });
  txt(s, "All 12 combinations were tried (36 s); the best 6 and the worst are shown. Lower CV MAE is better. The smallest model wins: 200 trees, depth 2, ν = 0.02. More trees, a higher rate or deeper trees all made validation error worse — a sign of over-fitting.",
    { x: 7.1, y: 5.1, w: 5.63, h: 1.9, fontSize: 14 });
  s.addNotes("CV MAE is measured on the log-ratio target, so it is not in orders. The ranking shows that more trees, a higher learning rate and deeper trees all make validation error worse — signs of over-fitting.");

  // ================================================================ 11 learning curve
  s = content("More trees is not always better", "Our model");
  img(s, IMG("06_boosting", "importance_and_learning_curve.png"), M, 1.4, 8.1, 2.386, "importance chart");
  stat(s, 9.2, 1.45, 3.5, "133", "trees give the lowest validation error (MAE 4.70 orders)");
  stat(s, 9.2, 3.0, 3.5, "4.87", "validation MAE at 800 trees: the extra trees only fit noise", C.accent2);
  card(s, 9.2, 4.55, 3.53, 2.45, "definition card");
  txt(s, [
    { text: "Over-fitting", options: { bold: true, fontSize: 16, color: HEX.dk2, breakLine: true } },
    { text: "means learning the noise of the training weeks instead of the real pattern: training error falls, but error on new weeks rises." }],
    { x: 9.45, y: 4.7, w: 3.1, h: 2.2, fontSize: 15 });
  bullets(s, [
    "Left — feature importance: festival days (0.34) and the recent level roll4 (0.21) drive most of the forecast",
    "A permutation test on unseen weeks confirms only these two; the other splits mostly fit training noise",
    "Right — learning curve: training error (blue) keeps falling, but validation error (orange) is lowest near 133 trees and then rises. 200 trees with a small rate stays close to that best point",
  ], { x: M, y: 4.85, w: 8.1, h: 2.15, fontSize: 14 });
  s.addNotes("Learning curve from staged_predict: validation = last 13 training weeks. Permutation importance = how much the error grows when one feature's values are shuffled.");

  // ================================================================ 12 results
  pres.addSection({ title: "Results" });
  s = content("Growth-free boosting is the most accurate forecaster", "Results");
  s.addChart(pres.charts.BAR, [{ name: "Test MAE",
      labels: ["Naive", "Boosting (original features)", "Bayesian", "Ridge", "Linear", "Boosting (growth-free)"],
      values: [6.761, 5.917, 5.356, 5.353, 5.353, 5.315] }],
    Object.assign({ x: M, y: 1.45, w: 7.7, h: 5.5, barDir: "bar", barGapWidthPct: 45,
      chartColors: [HEX.accent6, HEX.accent6, HEX.accent1, HEX.accent1, HEX.accent1, HEX.accent2],
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00", valAxisMinVal: 0, valAxisMaxVal: 8,
      showLegend: false, showTitle: true, title: "Test MAE on the same 1,508 store-weeks (lower is better)", objectName: "comparison chart" }, chartText));
  const st12 = (y, v, l, col) => {
    txt(s, v, { x: 8.75, y, w: 4.0, h: 0.7, fontSize: 36, bold: true, color: col || C.accent1, fontFace: THEME.headFontFace });
    txt(s, l, { x: 8.75, y: y + 0.72, w: 4.0, h: 0.75, fontSize: 14 });
  };
  st12(1.45, "5.32", "orders: the lowest error of all six models (R² 0.784, WAPE 16.3%)", C.accent2);
  st12(2.95, "−10%", "error versus plain boosting (5.92 → 5.32), thanks to the growth-free target");
  st12(4.45, "4.34", "noise floor: the best MAE any model could reach, because weekly orders are partly random");
  txt(s, "The gain over linear regression is small (0.7%): most of the signal is simple and linear, and all good models are close to the noise floor.",
    { x: 8.75, y: 6.0, w: 4.0, h: 1.0, fontSize: 14, italic: true, color: C.accent1 });
  s.addNotes("MAE = mean absolute error in orders per store-week. R² = share of the variation explained. WAPE = total absolute error as a share of total orders. The noise floor is the MAE of a perfect forecaster facing Poisson-like random variation.");

  // ================================================================ 13 code
  s = content("Implementation in scikit-learn", "Results");
  card(s, M, 1.4, 8.1, 3.75, "code card", "F4F6F5");
  txt(s, [
    "# 1. relative features",
    'REL = ["r1", "r2", "r4", "roll4", "store_age_w",',
    '       "festival_days", "monsoon", "woy", "area_*"]',
    "# 2. grid search with time-ordered CV",
    'grid = {"n_estimators": [200, 400, 800],',
    '        "learning_rate": [0.02, 0.05], "max_depth": [2, 3]}',
    "gb = GridSearchCV(",
    "    GradientBoostingRegressor(subsample=0.8, random_state=42),",
    "    grid, cv=TimeSeriesSplit(n_splits=5),",
    '    scoring="neg_mean_absolute_error")',
    "# 3. train on the growth-free target, convert back",
    "z_train = np.log(train.orders / train.roll4)",
    "gb.fit(train[REL], z_train)",
    "pred = test.roll4 * np.exp(gb.predict(test[REL]))",
  ].join("\n"), { x: M + 0.3, y: 1.55, w: 7.6, h: 3.5, fontSize: 13, fontFace: "Courier New", color: C.text1 });
  card(s, M, 5.3, 8.1, 1.7, "output card", C.text2);
  txt(s, "Output", { x: M + 0.3, y: 5.42, w: 3, h: 0.4, fontSize: 15, bold: true, color: C.accent3 });
  txt(s, "best_params_: n_estimators=200, learning_rate=0.02, max_depth=2\ntest MAE 5.315 · RMSE 6.94 · R² 0.784 · WAPE 16.3%",
    { x: M + 0.3, y: 5.85, w: 7.6, h: 0.95, fontSize: 14, fontFace: "Courier New", color: C.background1 });
  card(s, 9.0, 1.4, 3.73, 5.6, "steps card");
  txt(s, "What the code does", { x: 9.25, y: 1.6, w: 3.3, h: 0.45, fontSize: 17, bold: true, color: C.text2 });
  bullets(s, ["Lists the relative features (area_* = the three location columns)",
              "Tries all 12 combinations of trees, learning rate and depth with time-ordered CV",
              "Trains on the growth-free target z = log(orders ÷ roll4)",
              "Turns the predicted change back into orders: roll4 × e^prediction",
              "Scores the 26 test weeks with MAE, RMSE, R² and WAPE"], { x: 9.25, y: 2.1, w: 3.3, h: 4.8, fontSize: 15 });
  s.addNotes("Full code: notebook 06_PBL3_Gradient_Boosting.ipynb in the project folder. GridSearchCV refits the best setting on all training weeks after the search.");

  // ================================================================ 14 conclusion
  s = content("Conclusion", "Results", "DARK");
  const fin = [["Accuracy", "Growth-free gradient boosting is our most accurate model: MAE 5.32 orders and R² 0.784, only about one order above the noise floor."],
               ["Key insight", "Trees cannot forecast beyond values they have seen. Predicting the change from the recent level fixed this; without it boosting lost to linear regression."],
               ["Trade-offs", "Harder to explain than a regression formula, slower to tune, and no prediction interval by default (Bayesian regression gives one)."],
               ["Recommendation", "Use boosting when richer features arrive — prices, promotions, weather, demand regimes — where its non-linear power can really help."]];
  fin.forEach((c, i) => {
    const x = M + (i % 2) * 6.15, y = 1.45 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 5.95, 2.35, "conclusion card", C.accent1);
    txt(s, c[0], { x: x + 0.35, y: y + 0.25, w: 5.3, h: 0.5, fontSize: 20, bold: true, color: i === 3 ? C.accent3 : C.background1 });
    txt(s, c[1], { x: x + 0.35, y: y + 0.85, w: 5.3, h: 1.35, fontSize: 16, color: C.background2 });
  });
  txt(s, "Take-away: better forecasts now need new information, not a more complex model.",
    { x: M, y: 6.6, w: 12.1, h: 0.5, fontSize: 17, italic: true, bold: true, color: C.accent3 });
  s.addNotes("Gradient boosting confirmed that the regression baseline was already close to the limit of the available features; better forecasts now need new information rather than a more complex model.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();

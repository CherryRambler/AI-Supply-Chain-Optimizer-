"""Workflow diagrams for the report (sections 2.5, 3.5, 3.10, 4.4)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

plt.rcParams["font.family"] = "Liberation Serif"
FILL = {"input": "#E8ECF2", "prep": "#DCE8F7", "model": "#D6F0E4", "eval": "#FCE3D6", "out": "#F6EBC8"}
EDGE = {"input": "#5F6774", "prep": "#2A78D6", "model": "#1BAF7A", "eval": "#EB6834", "out": "#B98900"}


def flow(stages, path, cols=3, w=1.98, h=0.92, gx=0.3, gy=0.34):
    rows = (len(stages) + cols - 1) // cols
    W, H = cols * w + (cols - 1) * gx, rows * h + (rows - 1) * gy
    fig, ax = plt.subplots(figsize=(W + 0.1, H + 0.1))
    ax.set_xlim(-0.05, W + 0.05); ax.set_ylim(-0.05, H + 0.05); ax.axis("off")
    pos = []
    for i, (kind, title, body) in enumerate(stages):
        r, c = divmod(i, cols)
        if r % 2 == 1:
            c = cols - 1 - c                      # snake layout
        x, y = c * (w + gx), H - (r + 1) * h - r * gy
        pos.append((x, y))
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.01,rounding_size=0.07",
                                    fc=FILL[kind], ec=EDGE[kind], lw=1.1))
        ax.text(x + w / 2, y + h - 0.19, f"{i + 1}. {title}", ha="center", va="center", fontsize=8.6, weight="bold", color="#14213D")
        ax.text(x + w / 2, y + (h - 0.3) / 2, body, ha="center", va="center", fontsize=7.4, color="#2B3A55", linespacing=1.2)
    for i in range(len(stages) - 1):
        (x1, y1), (x2, y2) = pos[i], pos[i + 1]
        if abs(y1 - y2) < 1e-6:                   # same row
            start = (x1 + w, y1 + h / 2) if x2 > x1 else (x1, y1 + h / 2)
            end = (x2, y2 + h / 2) if x2 > x1 else (x2 + w, y2 + h / 2)
        else:                                     # next row
            start, end = (x1 + w / 2, y1), (x2 + w / 2, y2 + h)
        ax.add_patch(FancyArrowPatch(start, end, arrowstyle="-|>", mutation_scale=10, lw=1.1, color="#3A4A63"))
    fig.savefig(path, dpi=260, bbox_inches="tight", facecolor="white")
    plt.close(fig)


flow([
    ("input", "Raw dataset", "Kaggle CSV\n1,95,148 orders × 39 columns"),
    ("prep", "Data cleaning", "duplicates, formula check,\nstructural missing values"),
    ("prep", "Food filter", "7 food categories kept\n1,17,077 orders"),
    ("prep", "IPL-effect removal", "per-store excess evening\norders removed → 1,13,390"),
    ("prep", "Feature engineering", "store-week table (4,039 rows):\nlags, roll4, festival, area"),
    ("prep", "Time split + scaling", "train ≤ 23-03-2026, test 26 wks;\nz-score fitted on train"),
    ("model", "Regression models", "Naive · Linear (OLS)\nRidge (TS-CV) · Bayesian"),
    ("eval", "Evaluation", "MAE, RMSE, R², WAPE;\n95% interval coverage"),
    ("out", "Output", "weekly demand forecast per\ndark store → replenishment"),
], "fig_workflow_ch2.png")

flow([
    ("input", "Clean food orders", "1,13,390 orders\n(from Chapter 2)"),
    ("prep", "Store feature table", "58 stores × 20 candidate\nfeatures"),
    ("prep", "Feature selection", "signal-to-noise test →\n10 demand features"),
    ("prep", "Standardisation", "z-score of every feature"),
    ("model", "Choose k", "k = 2…8: elbow, silhouette,\nDB, CH → k = 3"),
    ("model", "Final K-Means", "k = 3, 50 random starts,\nseed 42"),
    ("eval", "Validation", "Ward, GMM, ARI stability,\nper-store silhouette"),
    ("eval", "Interpretation", "PCA plot, profiles, basket mix\nand hourly pattern"),
    ("out", "Output", "store_clusters.csv →\nstocking & staffing plans"),
], "fig_workflow_unit3.png")

flow([
    ("input", "Clean food orders", "1,13,390 orders\n(from Chapter 2)"),
    ("prep", "Daily sequence", "638 consecutive days\n(01-01-2025 to 30-09-2026)"),
    ("prep", "Observations", "demand index (growth and\nweekday removed), late rate"),
    ("model", "Train Gaussian HMMs", "2–5 states, Baum–Welch (EM),\n20 random restarts"),
    ("model", "Model selection", "lowest BIC → 3 hidden states"),
    ("model", "Parameters", "π, transition matrix A,\nGaussian emissions B"),
    ("model", "Decoding", "Viterbi state sequence +\nposterior probabilities"),
    ("eval", "Validation", "festival recall, 2025 → 2026\ntest, vs K-Means / GMM"),
    ("out", "Output", "daily_regimes.csv +\nnext-day regime forecast"),
], "fig_workflow_unit4.png")

flow([
    ("input", "Store-week dataset", "4,039 rows from Chapter 2\n(same features)"),
    ("prep", "Same time split", "train ≤ 23-03-2026,\ntest 26 weeks (as PBL 1)"),
    ("prep", "Growth-free target", "log(orders / roll4) +\nrelative lag features"),
    ("model", "Gradient boosting", "regression trees (CART)\nadded stage-wise"),
    ("model", "Hyperparameter search", "GridSearchCV, 5-fold\ntime-series CV (12 settings)"),
    ("model", "Final model", "200 trees, depth 2,\nlearning rate 0.02"),
    ("eval", "Evaluation", "MAE, RMSE, R², WAPE;\nimportance, learning curve"),
    ("eval", "Comparison", "same test weeks vs Naive,\nLinear, Ridge, Bayesian"),
    ("out", "Output", "weekly forecasts and\nmodel recommendation"),
], "fig_workflow_unit5.png")
print("diagrams written")

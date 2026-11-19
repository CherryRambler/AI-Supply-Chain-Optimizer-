"""Make every notebook chart readable at A4 text width (≈5.8 in): smaller figure sizes, 2-row layouts
for 3-panel charts, higher export resolution. Numbers and models are unchanged."""
import nbformat, sys, os

PROJ = sys.argv[1]
SETUP_OLD = '"axes.titleweight": "bold", "axes.titlesize": 12,'
SETUP_NEW = ('"axes.titleweight": "bold", "axes.titlesize": 9.5, "font.size": 8.5, "axes.labelsize": 8.5,\n'
             '    "xtick.labelsize": 7.5, "ytick.labelsize": 7.5, "legend.fontsize": 7.5,')

EDITS = {
"01_Create_Dataset.ipynb": [
    ('fig, axes = plt.subplots(1, 3, figsize=(15, 4))\nd_all',
     'fig = plt.figure(figsize=(6.5, 4.6)); gs = fig.add_gridspec(2, 2, height_ratios=[1, 1.2])\n'
     'axes = [fig.add_subplot(gs[0, :]), fig.add_subplot(gs[1, 0]), fig.add_subplot(gs[1, 1])]\nd_all'),
    ('axes[1].tick_params(axis="y", labelsize=8)', 'axes[1].tick_params(axis="y", labelsize=6.5)'),
    ('fig, ax = plt.subplots(figsize=(7, 3.6))', 'fig, ax = plt.subplots(figsize=(5.2, 2.8))'),
    ('f"{vb[i]:.1f}%", ha="center", fontsize=9)', 'f"{vb[i]:.1f}%", ha="center", fontsize=7.5)'),
    ('f"{va[i]:.1f}%", ha="center", fontsize=9)', 'f"{va[i]:.1f}%", ha="center", fontsize=7.5)'),
    ('fig, axes = plt.subplots(2, 1, figsize=(13, 5.6), sharex=True)', 'fig, axes = plt.subplots(2, 1, figsize=(6.5, 3.8), sharex=True)'),
],
"02_Clustering_KMeans.ipynb": [
    ('fig, ax = plt.subplots(figsize=(8.5, 6))', 'fig, ax = plt.subplots(figsize=(6.2, 3.9))'),
    ('ax.grid(axis="y", visible=False); ax.tick_params(axis="y", labelsize=8)', 'ax.grid(axis="y", visible=False); ax.tick_params(axis="y", labelsize=7)'),
    ('fig, axes = plt.subplots(1, 2, figsize=(12, 3.8))', 'fig, axes = plt.subplots(1, 2, figsize=(6.5, 2.5))'),
    ('fig, axes = plt.subplots(1, 2, figsize=(14, 5))', 'fig, axes = plt.subplots(1, 2, figsize=(6.6, 3.5), gridspec_kw={"width_ratios": [1, 1.05]})'),
    ('axes[0].scatter(P[m, 0], P[m, 1], s=60,', 'axes[0].scatter(P[m, 0], P[m, 1], s=22,'),
    ('axes[0].scatter(C[i, 0], C[i, 1], s=220,', 'axes[0].scatter(C[i, 0], C[i, 1], s=90,'),
    ('axes[0].legend(fontsize=8)\nz = ', 'axes[0].legend(fontsize=6.5, loc="lower left")\nz = '),
    ('axes[1].set_xticks(range(K)); axes[1].set_xticklabels([n.replace(" ", "\\n", 1) for n in ORDER], fontsize=8)',
     'axes[1].set_xticks(range(K)); axes[1].set_xticklabels([n.replace(" ", "\\n", 1) for n in ORDER], fontsize=6.5)'),
    ('axes[1].set_yticks(range(len(CLUSTER_FEATURES))); axes[1].set_yticklabels(CLUSTER_FEATURES, fontsize=8)',
     'axes[1].set_yticks(range(len(CLUSTER_FEATURES))); axes[1].set_yticklabels(CLUSTER_FEATURES, fontsize=6.5)'),
    ('f"{z.T.values[i, j]:+.1f}", ha="center", va="center", fontsize=8, color=INK)', 'f"{z.T.values[i, j]:+.1f}", ha="center", va="center", fontsize=6.5, color=INK)'),
    ('fig.colorbar(im, ax=axes[1], shrink=0.8, label="standardised value (0 = average store)")', 'fig.colorbar(im, ax=axes[1], shrink=0.8, label="z-score (0 = average store)")'),
    ('fig, axes = plt.subplots(1, 2, figsize=(14, 4.2))\nxpos', 'fig, axes = plt.subplots(1, 2, figsize=(6.5, 2.8))\nxpos'),
    ('fontsize=8, rotation=15)', 'fontsize=6.5, rotation=20)'),
    ('axes[0].set(title="Food categories in the basket", ylabel="% of orders containing it"); axes[0].legend(fontsize=8)',
     'axes[0].set(title="Food categories in the basket", ylabel="% of orders containing it"); axes[0].legend(fontsize=6)'),
    ('axes[1].set(title="When orders arrive", xlabel="hour of day", ylabel="% of the cluster\'s orders"); axes[1].legend(fontsize=8)',
     'axes[1].set(title="When orders arrive", xlabel="hour of day", ylabel="% of the cluster\'s orders"); axes[1].legend(fontsize=6)'),
    ('fig, ax = plt.subplots(figsize=(8, 4.5))', 'fig, ax = plt.subplots(figsize=(5.6, 2.9))'),
    ('ax.text(-0.05, y0 + len(v) / 2, name, ha="right", va="center", fontsize=8)', 'ax.text(-0.05, y0 + len(v) / 2, name, ha="right", va="center", fontsize=7)'),
],
"03_HMM.ipynb": [
    ('fig, ax = plt.subplots(figsize=(6.5, 3.4))', 'fig, ax = plt.subplots(figsize=(4.6, 2.5))'),
    ('fig, axes = plt.subplots(1, 2, figsize=(13, 4))', 'fig, axes = plt.subplots(1, 2, figsize=(6.5, 2.8))'),
    ('axes[0].set_xticks(range(N_STATES)); axes[0].set_xticklabels(SORDER, fontsize=8)', 'axes[0].set_xticks(range(N_STATES)); axes[0].set_xticklabels([s.replace(" / ", "/\\n") for s in SORDER], fontsize=6.5)'),
    ('axes[0].set_yticks(range(N_STATES)); axes[0].set_yticklabels(SORDER, fontsize=8)', 'axes[0].set_yticks(range(N_STATES)); axes[0].set_yticklabels([s.replace(" / ", "/\\n") for s in SORDER], fontsize=6.5)'),
    ('f"{trans.values[i, j]:.3f}", ha="center", va="center", fontsize=9,', 'f"{trans.values[i, j]:.3f}", ha="center", va="center", fontsize=7.5,'),
    ('fig, axes = plt.subplots(2, 1, figsize=(14, 6), sharex=True)', 'fig, axes = plt.subplots(2, 1, figsize=(6.5, 4.2), sharex=True)'),
    ('axes[0].scatter(daily.index[m], daily.demand_index[m], s=10,', 'axes[0].scatter(daily.index[m], daily.demand_index[m], s=4,'),
    ('axes[1].scatter(daily.index[m], daily.late_rate[m] * 100, s=10,', 'axes[1].scatter(daily.index[m], daily.late_rate[m] * 100, s=4,'),
    ('axes[0].legend(fontsize=8, ncol=3)', 'axes[0].legend(fontsize=7, ncol=3, markerscale=2)'),
    ('textcoords="offset points", fontsize=7, ha="center", color=INK2)', 'textcoords="offset points", fontsize=6, ha="center", color=INK2)'),
],
"04_Comparison_and_Results.ipynb": [
    ('fig, axes = plt.subplots(1, 3, figsize=(15, 3.8))', 'fig, axes = plt.subplots(1, 3, figsize=(6.5, 2.4))'),
    ('fmt.format(v), ha="center", va="bottom", fontsize=9)', 'fmt.format(v), ha="center", va="bottom", fontsize=7)'),
    ('ax.tick_params(axis="x", labelsize=8)', 'ax.tick_params(axis="x", labelsize=5.5)'),
    ('fig, ax = plt.subplots(figsize=(14, 2.8))', 'fig, ax = plt.subplots(figsize=(6.5, 2.1))'),
    ('labels=SORDER, fontsize=8, ncol=3,', 'labels=SORDER, fontsize=7, ncol=3,'),
    ('ax.set_title("Label of each day: methods without time order flicker; the HMM gives stable regimes")',
     'ax.set_title("Label of each day: methods without time order flicker, the HMM is stable")'),
    ('fig, ax = plt.subplots(figsize=(8, 3.6))', 'fig, ax = plt.subplots(figsize=(5.6, 2.5))'),
    ('ax.set_xticks(range(3)); ax.set_xticklabels(SORDER, fontsize=8); ax.set_yticks(range(3)); ax.set_yticklabels(ORDER, fontsize=8)',
     'ax.set_xticks(range(3)); ax.set_xticklabels(SORDER, fontsize=7.5); ax.set_yticks(range(3)); ax.set_yticklabels(ORDER, fontsize=7.5)'),
    ('f"{tab.values[i, j]:.1f}%", ha="center", va="center", fontsize=9)', 'f"{tab.values[i, j]:.1f}%", ha="center", va="center", fontsize=8)'),
],
"05_PBL1_Regression.ipynb": [
    ('fig, axes = plt.subplots(1, 3, figsize=(16, 4))\naxes[0].hist',
     'fig = plt.figure(figsize=(6.5, 4.6)); gs = fig.add_gridspec(2, 2, height_ratios=[1, 1.1])\n'
     'axes = [fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[1, :]), fig.add_subplot(gs[0, 1])]\naxes[0].hist'),
    ('fig, axes = plt.subplots(1, 2, figsize=(14, 4.2))\ny_pos', 'fig, axes = plt.subplots(1, 2, figsize=(6.6, 3.0))\ny_pos'),
    ('axes[0].set_yticks(y_pos); axes[0].set_yticklabels(FEATURES, fontsize=8)', 'axes[0].set_yticks(y_pos); axes[0].set_yticklabels(FEATURES, fontsize=6.5)'),
    ('axes[0].set(title="Model coefficients (standardised features)", xlabel="weekly orders per 1 std of the feature")',
     'axes[0].set(title="Coefficients (standardised features)", xlabel="weekly orders per 1 std")'),
    ('axes[0].legend(fontsize=8); axes[0].grid(axis="y", visible=False)\nfor j', 'axes[0].legend(fontsize=6.5); axes[0].grid(axis="y", visible=False)\nfor j'),
    ('axes[1].set_xscale("log"); axes[1].legend(fontsize=7, ncol=2)', 'axes[1].set_xscale("log"); axes[1].legend(fontsize=6, ncol=2)'),
    ('fig, axes = plt.subplots(1, 3, figsize=(16, 4.3))\naxes[0].scatter(y_test, best, s=8,',
     'fig = plt.figure(figsize=(6.5, 4.8)); gs = fig.add_gridspec(2, 2, height_ratios=[1.15, 1])\n'
     'axes = [fig.add_subplot(gs[0, 0]), fig.add_subplot(gs[1, :]), fig.add_subplot(gs[0, 1])]\naxes[0].scatter(y_test, best, s=4,'),
    ('title=f"Bayesian forecast with uncertainty — store {s_id}"', 'title=f"Bayesian 95% interval — store {s_id}"'),
    ('axes[0].set(title="Actual vs predicted (Linear Regression)",', 'axes[0].set(title="Actual vs predicted (linear)",'),
],
"06_PBL3_Gradient_Boosting.ipynb": [
    ('fig, axes = plt.subplots(1, 2, figsize=(13, 4))', 'fig, axes = plt.subplots(1, 2, figsize=(6.5, 2.8))'),
    ('axes[0].scatter(tr_max, test_mean.reindex(tr_max.index), s=30, color=BLUE)', 'axes[0].scatter(tr_max, test_mean.reindex(tr_max.index), s=12, color=BLUE)'),
    ('axes[0].set(title="Store level: training maximum vs test average", xlabel="highest weekly orders in training",\n            ylabel="average weekly orders in test",',
     'axes[0].set(title="Per store: training max vs test mean", xlabel="highest weekly orders in training",\n            ylabel="mean weekly orders in test",'),
    ('fig, axes = plt.subplots(1, 2, figsize=(14, 4.2))', 'fig, axes = plt.subplots(1, 2, figsize=(6.6, 2.8))'),
    ('axes[0].tick_params(axis="y", labelsize=8)', 'axes[0].tick_params(axis="y", labelsize=6.5)'),
    ('axes[0].set(title="Feature importance (impurity reduction)", xlabel="share of total importance")',
     'axes[0].set(title="Feature importance (impurity)", xlabel="share of total importance")'),
    ('label="validation (last 13 training weeks)")', 'label="validation (last 13 train weeks)")'),
    ('ylabel="MAE (weekly orders)"); axes[1].legend(fontsize=8)', 'ylabel="MAE (weekly orders)"); axes[1].legend(fontsize=6.5)'),
    ('fig, ax = plt.subplots(figsize=(10, 4.2))', 'fig, ax = plt.subplots(figsize=(6.3, 2.7))'),
    ('ax.text(v + 0.05, i, f"{v:.2f}", va="center", fontsize=9)', 'ax.text(v + 0.05, i, f"{v:.2f}", va="center", fontsize=7.5)'),
    ('"noise floor", fontsize=8, color=INK2)', '"noise floor", fontsize=7, color=INK2)'),
    ('ax.set(title="Test MAE (weekly orders per store) — lower is better", xlabel="MAE")', 'ax.set(title="Test MAE per store-week (lower is better)", xlabel="MAE (orders)")'),
],
}

for name, edits in EDITS.items():
    path = os.path.join(PROJ, name)
    nb = nbformat.read(path, 4)
    src = [c.source for c in nb.cells]
    done = 0
    for c in nb.cells:
        if c.cell_type != "code":
            continue
        s = c.source
        if SETUP_OLD in s:
            s = s.replace(SETUP_OLD, SETUP_NEW); done += 1
        s = s.replace('"lines.linewidth": 2,', '"lines.linewidth": 1.5,').replace("dpi=150, bbox_inches", "dpi=240, bbox_inches")
        c.source = s
    whole = "\n\x00\n".join(c.source for c in nb.cells)
    for old, new in edits:
        n = sum(c.source.count(old) for c in nb.cells if c.cell_type == "code")
        assert n == 1, (name, old[:70], n)
        for c in nb.cells:
            if c.cell_type == "code" and old in c.source:
                c.source = c.source.replace(old, new)
    assert done == 1, (name, done)
    nbformat.write(nb, path)
    print(name, "ok", len(edits), "edits")

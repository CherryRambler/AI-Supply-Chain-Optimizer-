# Extra chart for the gradient boosting deck (errors slide). Run from the project folder:
#   python presentation/make_deck_figures.py
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

BLUE, ORANGE, AQUA, INK, INK2, MUTED, GRID = "#2a78d6", "#eb6834", "#1baf7a", "#0b0b0b", "#52514e", "#898781", "#e6e5e0"
plt.rcParams.update({
    "figure.facecolor": "#fcfcfb", "axes.facecolor": "#fcfcfb", "axes.edgecolor": MUTED, "axes.labelcolor": INK2,
    "axes.titlecolor": INK, "axes.titleweight": "bold", "axes.titlesize": 9.5, "font.size": 8.5, "axes.labelsize": 8.5,
    "xtick.labelsize": 7.5, "ytick.labelsize": 7.5, "legend.fontsize": 7.5, "axes.grid": True, "grid.color": GRID,
    "axes.spines.top": False, "axes.spines.right": False, "legend.frameon": False})

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
OUT = os.path.join(ROOT, "presentation", "assets")
os.makedirs(OUT, exist_ok=True)

gb = pd.read_csv(os.path.join(ROOT, "Unit_5_PBL3_Gradient_Boosting", "results", "06_boosting", "test_predictions.csv"), parse_dates=["week"])
lin = pd.read_csv(os.path.join(ROOT, "Unit_1_2_PBL1_Data_and_Regression", "results", "05_regression", "test_predictions.csv"), parse_dates=["week"])
wk = pd.read_csv(os.path.join(ROOT, "dataset", "store_weekly.csv"), parse_dates=["week"])
p = (gb.merge(lin[["store_id", "week", "Linear Regression"]], on=["store_id", "week"])
       .merge(wk[["store_id", "week", "festival_days"]], on=["store_id", "week"]))
fest = p.festival_days > 0

fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.9), gridspec_kw={"width_ratios": [1, 1.35]})
ax[0].scatter(p.actual[~fest], p.gb_growth_free[~fest], s=6, alpha=0.45, color=BLUE, label="normal weeks")
ax[0].scatter(p.actual[fest], p.gb_growth_free[fest], s=12, alpha=0.9, color=ORANGE, label="festival weeks")
lim = [0, max(p.actual.max(), p.gb_growth_free.max()) * 1.05]
ax[0].plot(lim, lim, color=INK2, lw=1)
ax[0].set(title="Actual vs predicted (gradient boosting)", xlabel="actual weekly orders", ylabel="predicted", xlim=lim, ylim=lim)
ax[0].legend(loc="upper left", markerscale=1.5)

t = p.groupby("week")[["actual", "gb_growth_free", "Linear Regression", "gb_original"]].sum()
ax[1].plot(t.index, t.actual, color=INK, lw=2.2, label="actual")
ax[1].plot(t.index, t.gb_growth_free, color=ORANGE, lw=1.6, label="boosting (growth-free)")
ax[1].plot(t.index, t["Linear Regression"], color=BLUE, lw=1.3, label="linear regression")
ax[1].plot(t.index, t.gb_original, color=MUTED, lw=1.2, ls="--", label="boosting (original features)")
ax[1].set(title="Network total per week (26 test weeks)", ylabel="orders")
ax[1].xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
ax[1].legend(loc="upper left")
plt.tight_layout()
fig.savefig(os.path.join(OUT, "boosting_errors.png"), dpi=220)
print("saved", os.path.join(OUT, "boosting_errors.png"))

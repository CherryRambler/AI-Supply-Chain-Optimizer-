"""AI Supply Chain Optimizer - Indian Quick-Commerce Food Demand.

Streamlit front-end for the trained models of the project. The app only LOADS the saved models and
result files; it never trains or overwrites anything.

Run:  streamlit run app.py
"""
from __future__ import annotations

import json
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------------------
# Paths (relative to the project root = the folder that contains this file)
# --------------------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent
U1 = ROOT / "Unit_1_2_PBL1_Data_and_Regression" / "results" / "05_regression"
U3 = ROOT / "Unit_3_PBL2_Clustering_KMeans" / "results" / "02_clustering"
U4 = ROOT / "Unit_4_PBL2_HMM" / "results" / "03_hmm"
U5 = ROOT / "Unit_5_PBL3_Gradient_Boosting" / "results" / "06_boosting"
WEEKLY_CSV = ROOT / "dataset" / "store_weekly.csv"

GB_MODEL = U5 / "gradient_boosting.joblib"
REG_MODEL = U1 / "regression_models.joblib"
KM_MODEL = U3 / "kmeans_model.joblib"
HMM_MODEL = U4 / "hmm_model.joblib"

# --------------------------------------------------------------------------------------
# One consistent colour palette (same as the notebooks' charts)
# --------------------------------------------------------------------------------------
BLUE, ORANGE, AQUA, YELLOW = "#2a78d6", "#eb6834", "#1baf7a", "#eda100"
NAVY, INK, MUTED, GRID = "#12305c", "#1f2933", "#898781", "#e6e5e0"
REGIME_COLORS = {"Normal": BLUE, "Monsoon / disrupted": AQUA, "Festival peak": ORANGE}
REGIME_ORDER = list(REGIME_COLORS)
MODEL_COLORS = {"Naive (last week)": MUTED, "Linear Regression": BLUE, "Ridge Regression": "#6aa5e8",
                "Bayesian Regression": AQUA, "Gradient Boosting": ORANGE}

GB_KEY = "Gradient Boosting (growth-free target)"   # row name used in summary.json / comparison csv

st.set_page_config(page_title="AI Supply Chain Optimizer", page_icon="🛵", layout="wide")

st.markdown(
    f"""
    <style>
      #MainMenu, footer {{ visibility: hidden; }}
      .block-container {{ padding-top: 1.4rem; max-width: 1250px; }}
      h5 {{ margin-top: .4rem; padding-bottom: .25rem; border-bottom: 2px solid rgba(42,120,214,.25); }}
      .app-header {{ background: linear-gradient(110deg, {NAVY} 0%, #1d4f9c 55%, {BLUE} 100%); color: #fff;
                    padding: 1.3rem 1.8rem; border-radius: 16px; margin-bottom: 1.4rem;
                    box-shadow: 0 6px 18px rgba(18,48,92,.28); }}
      .app-header .kicker {{ letter-spacing: .12em; font-size: .72rem; font-weight: 600; color: #a9c8f5;
                            text-transform: uppercase; }}
      .app-header h1 {{ color: #fff; font-size: 2rem; margin: .15rem 0 0 0; padding: 0; }}
      .app-header p  {{ color: #dbe8fb; margin: .3rem 0 0 0; font-size: .98rem; }}
      .pill {{ display: inline-block; background: rgba(255,255,255,.16); color: #fff; border-radius: 999px;
              padding: .12rem .7rem; font-size: .75rem; margin: .7rem .4rem 0 0; }}
      .hero {{ background: linear-gradient(160deg, rgba(235,104,52,.16), rgba(235,104,52,.04));
              border: 1px solid rgba(235,104,52,.35); border-radius: 16px; padding: 1.1rem 1.3rem; }}
      .hero .label {{ font-size: .85rem; color: {MUTED}; text-transform: uppercase; letter-spacing: .08em; }}
      .big-number {{ font-size: 4rem; font-weight: 800; color: {ORANGE}; line-height: 1.05; }}
      .big-caption {{ color: {MUTED}; font-size: .92rem; }}
      div[data-testid="stMetric"] {{ background: rgba(42,120,214,.10); border: 1px solid rgba(42,120,214,.22);
                                    border-left: 5px solid {BLUE}; padding: .65rem 1rem; border-radius: 10px; }}
      div[data-testid="stMetricValue"] {{ font-weight: 700; }}
      div[data-testid="stMetric"] {{ min-height: 6.2rem; }}
      section[data-testid="stSidebar"] {{ border-right: 1px solid rgba(42,120,214,.25); }}
      .side-brand {{ font-weight: 800; font-size: 1.15rem; color: {BLUE}; margin-bottom: .1rem; }}
      .side-sub {{ color: {MUTED}; font-size: .8rem; margin-bottom: .8rem; }}
      div[data-testid="stPlotlyChart"], div[data-testid="stDataFrame"] {{ border-radius: 10px; }}
    </style>
    """,
    unsafe_allow_html=True,
)

PAGE_SUBTITLES = {
    "Demand Forecast": "Weekly orders per dark store — gradient boosting on a growth-free target",
    "Store Types (K-Means)": "Unit 3 · K-Means groups the 58 stores into 3 types (PCA only draws the picture)",
    "Demand Regimes (HMM)": "Unit 4 · A Hidden Markov Model decodes Normal, Monsoon and Festival days",
    "Model Comparison": "Unit 5 · Naive vs regression vs gradient boosting on the held-out test weeks",
    "About": "Project, dataset, team and methods",
}


def header(_subtitle: str = "") -> None:
    page = st.session_state.get("nav", "Demand Forecast")
    st.markdown(
        f'<div class="app-header"><div class="kicker">AI Supply Chain Optimizer · Indian Quick-Commerce Food Demand</div>'
        f"<h1>{page}</h1><p>{PAGE_SUBTITLES.get(page, '')}</p>"
        '<span class="pill">58 dark stores</span><span class="pill">Kaggle · 2025-2026</span>'
        '<span class="pill">Saved models · no retraining</span></div>',
        unsafe_allow_html=True,
    )


def need(*paths: Path) -> None:
    """Show a clear message and stop the page if any required file is missing."""
    missing = [p for p in paths if not Path(p).exists()]
    if missing:
        st.error("Required project file(s) not found. Run the notebooks first or restore these files:\n\n"
                 + "\n".join(f"- `{Path(p).relative_to(ROOT)}`" for p in missing))
        st.stop()


# --------------------------------------------------------------------------------------
# Cached loaders: models with st.cache_resource, tables with st.cache_data
# --------------------------------------------------------------------------------------
@st.cache_resource(show_spinner="Loading model…")
def load_model(path: str):
    return joblib.load(path)


@st.cache_data(show_spinner=False)
def load_csv(path: str, **kw) -> pd.DataFrame:
    return pd.read_csv(path, **kw)


@st.cache_data(show_spinner=False)
def load_json(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data(show_spinner=False)
def load_weekly() -> pd.DataFrame:
    return pd.read_csv(WEEKLY_CSV, parse_dates=["week"]).sort_values(["store_id", "week"]).reset_index(drop=True)


def show_png(path: Path, caption: str) -> None:
    if path.exists():
        st.image(str(path), caption=caption, width="stretch")
    else:
        st.warning(f"Chart image not found: `{path.relative_to(ROOT)}`")


def style(fig: go.Figure, height: int = 380, **kw) -> go.Figure:
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=40, b=10), template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(color="#e6e8eb"), legend=dict(orientation="h", y=-0.2), **kw)
    return fig


# --------------------------------------------------------------------------------------
# Forecast logic - identical to the notebook 06_PBL3_Gradient_Boosting.ipynb
#   z = log(orders / roll4)     features r1 = lag1/roll4, r2 = lag2/roll4, r4 = lag4/roll4, woy
#   forecast = roll4 * exp(model.predict(features))
# --------------------------------------------------------------------------------------
def build_features(store_rows: pd.DataFrame, week: pd.Timestamp, lag1, lag2, lag3, lag4,
                   festival_days, monsoon) -> tuple[pd.DataFrame, pd.DataFrame, float]:
    """Return (boosting features, regression features, roll4) for one store-week."""
    roll4 = (lag1 + lag2 + lag3 + lag4) / 4.0           # roll4 = mean of the previous 4 weeks (notebook 05)
    last = store_rows.iloc[-1]
    store_age = int(last.store_age_w) + (week - last.week).days // 7   # store_age_w counts weeks since opening
    area = {c: int(last[c]) for c in ("area_Premium Residential", "area_Residential", "area_Student / PG Hub")}
    woy = int(week.isocalendar().week)
    base = {"roll4": roll4, "store_age_w": store_age, "festival_days": int(festival_days),
            "monsoon": int(monsoon), **area}
    gb_row = {"r1": lag1 / roll4, "r2": lag2 / roll4, "r4": lag4 / roll4, "woy": woy, **base}
    reg_row = {"lag1": lag1, "lag2": lag2, "lag4": lag4, **base}
    return pd.DataFrame([gb_row]), pd.DataFrame([reg_row]), roll4


def predict_all(gb_bundle: dict, reg_bundle: dict, gb_x: pd.DataFrame, reg_x: pd.DataFrame,
                roll4: float, lag1: float) -> dict[str, float]:
    z = float(gb_bundle["model"].predict(gb_x[gb_bundle["features"]])[0])
    gb = roll4 * float(np.exp(z))
    lin = float(reg_bundle["linear"].predict(reg_bundle["scaler"].transform(reg_x[reg_bundle["features"]]))[0])
    return {"Gradient Boosting": gb, "Linear Regression": lin, "Naive (last week)": float(lag1)}


def is_monsoon_month(d: pd.Timestamp) -> bool:
    return d.month in (6, 7, 8, 9, 10)                  # same rule as the dataset


# --------------------------------------------------------------------------------------
# Page 1 - Demand forecast
# --------------------------------------------------------------------------------------
def page_forecast() -> None:
    header("Weekly demand forecast per dark store (gradient boosting, growth-free target)")
    need(WEEKLY_CSV, GB_MODEL, REG_MODEL, U5 / "summary.json")
    weekly = load_weekly()
    gb_bundle, reg_bundle = load_model(str(GB_MODEL)), load_model(str(REG_MODEL))
    rmse = float(load_json(str(U5 / "summary.json"))[f"{GB_KEY}|RMSE"])

    stores = sorted(weekly.store_id.unique())
    fest_by_week = weekly.groupby("week").festival_days.max()

    c_store, c_mode = st.columns([1, 2])
    store = c_store.selectbox("Store", stores, index=0)
    mode = c_mode.radio("Input source", ["Enter values manually", "Load from data (real store-week)"],
                        horizontal=True)
    rows = weekly[weekly.store_id == store].reset_index(drop=True)
    last_week = rows.week.iloc[-1]
    from_data = mode.startswith("Load")
    actual = None

    if from_data:
        test_start = pd.Timestamp("2026-03-30")                     # start of the held-out test period
        choices = rows[rows.week >= test_start].week.dt.date.tolist() or rows.week.dt.date.tolist()
        pick = st.selectbox("Store-week (test period)", choices, index=len(choices) - 1)
        row = rows[rows.week == pd.Timestamp(pick)].iloc[0]
        week = pd.Timestamp(row.week)
        v1, v2, v4 = int(row.lag1), int(row.lag2), int(row.lag4)
        v3 = int(round(4 * row.roll4 - row.lag1 - row.lag2 - row.lag4))   # lag3 implied by roll4
        v_fest, v_mon = int(row.festival_days), bool(row.monsoon)
        actual = int(row.orders)
        ukey = f"data_{store}_{pick}"
    else:
        default_week = (last_week + pd.Timedelta(weeks=1)).date()
        picked = st.date_input("Week start date (weeks start on Monday)", value=default_week, key=f"d_{store}")
        picked = pd.Timestamp(picked)
        week = picked - pd.Timedelta(days=picked.weekday())         # snap to Monday
        if week != picked:
            st.info(f"Date moved to the Monday of that week: **{week:%d-%m-%Y}**.")
        tail = rows.orders.tolist()[-4:]                            # latest rows of the store
        v1, v2, v3, v4 = tail[-1], tail[-2], tail[-3], tail[-4]
        v_fest = int(fest_by_week.get(week, 0))
        v_mon = is_monsoon_month(week)
        ukey = f"man_{store}_{week:%Y%m%d}"

    st.markdown("##### Inputs")
    i1, i2, i3, i4, i5, i6 = st.columns([1, 1, 1, 1, 1, 1])
    lag1 = i1.number_input("Orders 1 week ago", value=v1, step=1, key=f"l1_{ukey}", disabled=from_data)
    lag2 = i2.number_input("Orders 2 weeks ago", value=v2, step=1, key=f"l2_{ukey}", disabled=from_data)
    lag3 = i3.number_input("Orders 3 weeks ago", value=v3, step=1, key=f"l3_{ukey}", disabled=from_data,
                           help="Needed to compute roll4 = average of the last 4 weeks.")
    lag4 = i4.number_input("Orders 4 weeks ago", value=v4, step=1, key=f"l4_{ukey}", disabled=from_data)
    fest = i5.number_input("Festival days in the week (0-7)", value=v_fest, step=1, key=f"f_{ukey}",
                           disabled=from_data)
    i6.write("")
    monsoon = i6.checkbox("Monsoon week", value=v_mon, key=f"m_{ukey}", disabled=from_data)

    # ---- validation: clear messages instead of crashes ----
    errors = []
    if min(lag1, lag2, lag3, lag4) < 0:
        errors.append("Orders cannot be negative. Please enter 0 or more for every past week.")
    if not 0 <= fest <= 7:
        errors.append("Festival days must be between 0 and 7.")
    if min(lag1, lag2, lag3, lag4) >= 0 and (lag1 + lag2 + lag3 + lag4) == 0:
        errors.append("All four past weeks are 0, so roll4 = 0 and the growth-free model cannot compute a ratio. "
                      "Enter at least one non-zero week.")
    age = int(rows.store_age_w.iloc[-1]) + (week - last_week).days // 7
    if age < 0:
        errors.append(f"{week:%d-%m-%Y} is before store {store} opened. Choose a later week.")
    if errors:
        for e in errors:
            st.error(e)
        return

    gb_x, reg_x, roll4 = build_features(rows, week, lag1, lag2, lag3, lag4, fest, monsoon)
    preds = predict_all(gb_bundle, reg_bundle, gb_x, reg_x, roll4, lag1)
    gb = preds["Gradient Boosting"]
    lo, hi = max(gb - rmse, 0.0), gb + rmse

    st.divider()
    left, right = st.columns([1, 2])
    with left:
        st.markdown(f"##### Predicted orders — {store}, week of {week:%d-%m-%Y}")
        st.markdown(f'<div class="hero"><div class="label">Predicted orders this week</div>'
                    f'<div class="big-number">{gb:,.0f}</div>'
                    f'<div class="big-caption">likely range ≈ {lo:,.0f} – {hi:,.0f} '
                    f"(± the model's test RMSE of {rmse:.1f})</div></div>", unsafe_allow_html=True)
        st.caption(f"roll4 = {roll4:.2f} · exact model output: {gb:.3f}")
        if actual is not None:
            m1, m2 = st.columns(2)
            m1.metric("Actual orders", f"{actual:,}")
            m2.metric("Prediction error", f"{gb - actual:+.1f}", help="predicted − actual")

    with right:
        hist = rows[rows.week < week] if from_data else rows
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=hist.week, y=hist.orders, mode="lines", name="Past weekly orders",
                                 line=dict(color=BLUE, width=2)))
        fig.add_trace(go.Scatter(x=[week], y=[gb], mode="markers", name="Forecast (gradient boosting)",
                                 marker=dict(color=ORANGE, size=13, symbol="diamond"),
                                 error_y=dict(type="data", symmetric=False, array=[hi - gb], arrayminus=[gb - lo],
                                              color=ORANGE, thickness=2)))
        if actual is not None:
            fig.add_trace(go.Scatter(x=[week], y=[actual], mode="markers", name="Actual",
                                     marker=dict(color=INK, size=11, symbol="circle-open", line=dict(width=2))))
        if not hist.empty:      # dotted connector from the last known week to the forecast
            fig.add_trace(go.Scatter(x=[hist.week.iloc[-1], week], y=[hist.orders.iloc[-1], gb], mode="lines",
                                     line=dict(color=ORANGE, dash="dot"), showlegend=False, hoverinfo="skip"))
        fig.update_layout(title=f"Weekly orders — store {store}", xaxis_title="Week start",
                          yaxis_title="Orders per week")
        st.plotly_chart(style(fig, 360), width="stretch")

    st.markdown("##### Comparison with simpler forecasts")
    cols = st.columns(3)
    for col, name in zip(cols, ["Gradient Boosting", "Linear Regression", "Naive (last week)"]):
        d = None if name == "Gradient Boosting" else f"{preds[name] - gb:+.1f} vs boosting"
        col.metric(name, f"{preds[name]:,.1f}", delta=d, delta_color="off")
    comp = pd.DataFrame({"Model": list(preds), "Forecast": list(preds.values())})
    if actual is not None:
        comp["Abs. error"] = (comp.Forecast - actual).abs()
        st.dataframe(comp.style.format({"Forecast": "{:.1f}", "Abs. error": "{:.1f}"}), hide_index=True,
                     width="stretch")
    bar = px.bar(comp, x="Model", y="Forecast", color="Model", color_discrete_map=MODEL_COLORS,
                 text=comp.Forecast.round(1))
    if actual is not None:
        bar.add_hline(y=actual, line_dash="dash", line_color=INK, annotation_text=f"actual = {actual}")
    bar.update_layout(showlegend=False, xaxis_title="Model", yaxis_title="Forecast orders for the week")
    st.plotly_chart(style(bar, 300), width="stretch")


# --------------------------------------------------------------------------------------
# Page 2 - Store types (K-Means)
# --------------------------------------------------------------------------------------
def page_clusters() -> None:
    header("Unit 3 — K-Means store types (PCA used only for the 2-D picture)")
    need(KM_MODEL, U3 / "store_clusters.csv", U3 / "cluster_profiles.csv", U3 / "summary.json")
    km = load_model(str(KM_MODEL))
    stores = load_csv(str(U3 / "store_clusters.csv"))
    profiles = load_csv(str(U3 / "cluster_profiles.csv"), index_col=0)
    summ = load_json(str(U3 / "summary.json"))

    store = st.selectbox("Store", stores.store_id.tolist())
    row = stores[stores.store_id == store].iloc[[0]]
    cid = int(km["kmeans"].predict(km["scaler"].transform(row[km["features"]]))[0])   # the saved model decides
    name = km["names"][cid]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Store type", name)
    c2.metric("City / area", f"{row.city.iloc[0]} · {row.area_type.iloc[0]}")
    c3.metric("Silhouette of this store", f"{row.silhouette.iloc[0]:.2f}",
              help="Closer to 1 = clearly inside its cluster.")
    c4.metric("Stores of this type", f"{summ['cluster_sizes'][name]} of {summ['n_stores']}")

    left, right = st.columns([1, 1.2])
    with left:
        st.markdown("##### Cluster profile (mean of each feature)")
        prof = profiles.copy()
        prof.columns = [f"{c} ★" if c == name else c for c in prof.columns]
        st.dataframe(prof.style.format("{:.2f}"), width="stretch", height=480)
        st.caption(f"★ = the type of store {store}.  k = {summ['k']}, silhouette {summ['silhouette']:.2f}, "
                   f"stability ARI {summ['stability_ari']:.2f}.")
    with right:
        show_png(U3 / "clusters_pca_and_profiles.png", "PCA view of the 58 stores and the cluster profiles")


# --------------------------------------------------------------------------------------
# Page 3 - Demand regimes (HMM)
# --------------------------------------------------------------------------------------
def page_regimes() -> None:
    header("Unit 4 — Hidden Markov Model: demand regimes")
    need(HMM_MODEL, U4 / "daily_regimes.csv", U4 / "regime_summary.csv", U4 / "next_day_forecast.csv")
    hmm = load_model(str(HMM_MODEL))
    daily = load_csv(str(U4 / "daily_regimes.csv"), parse_dates=["order_date"])
    summary = load_csv(str(U4 / "regime_summary.csv"), index_col=0)
    nxt = load_csv(str(U4 / "next_day_forecast.csv"), index_col=0)

    counts = daily.state.value_counts()
    cols = st.columns(3)
    for col, reg in zip(cols, REGIME_ORDER):
        col.metric(reg, f"{counts.get(reg, 0)} days", help=f"{counts.get(reg, 0) / len(daily):.0%} of all days")

    fig = go.Figure()
    for reg in REGIME_ORDER:
        d = daily[daily.state == reg]
        fig.add_trace(go.Bar(x=d.order_date, y=d.orders, name=reg, marker_color=REGIME_COLORS[reg],
                             hovertemplate="%{x|%d-%m-%Y}<br>%{y} orders<extra>" + reg + "</extra>"))
    fig.update_layout(title="Decoded regime of each day (bar height = orders that day)", barmode="overlay",
                      xaxis_title="Date", yaxis_title="Orders per day")
    st.plotly_chart(style(fig, 380), width="stretch")

    st.markdown("##### Look up a day")
    day = st.date_input("Date", value=daily.order_date.max().date(), min_value=daily.order_date.min().date(),
                        max_value=daily.order_date.max().date())
    r = daily[daily.order_date == pd.Timestamp(day)]
    if r.empty:
        st.warning("No data for that date.")
    else:
        r = r.iloc[0]
        a, b, c = st.columns(3)
        a.metric("Regime", r.state)
        b.metric("Confidence", f"{r.p_state:.0%}")
        c.metric("Orders / event", f"{int(r.orders)} · {r.event}")

    left, right = st.columns(2)
    with left:
        names = hmm["state_names"]                                       # model state index -> regime name
        order = [next(i for i, n in names.items() if n == reg) for reg in REGIME_ORDER]
        tm = np.asarray(hmm["hmm"].transmat_)[np.ix_(order, order)]      # reorder to Normal, Monsoon, Festival
        z = pd.DataFrame(tm, index=[f"from {r}" for r in REGIME_ORDER], columns=[f"to {r}" for r in REGIME_ORDER])
        heat = px.imshow(z, text_auto=".3f", color_continuous_scale=["#16222f", BLUE, "#8fc0ff"], zmin=0, zmax=1,
                         aspect="auto")
        heat.update_layout(title="Transition matrix (probability of tomorrow's regime)",
                           xaxis_title="Tomorrow's regime", yaxis_title="Today's regime",
                           coloraxis_colorbar=dict(title="P"))
        st.plotly_chart(style(heat, 360), width="stretch")
    with right:
        st.markdown("##### Next-day forecast")
        st.dataframe(nxt.style.format("{:.1%}"), width="stretch")
        for col in nxt.columns:
            st.caption(f"Most likely regime, {col}: **{nxt[col].idxmax()}** ({nxt[col].max():.1%})")
        st.markdown("##### Regime summary (daily averages)")
        st.dataframe(summary.style.format({"orders": "{:.1f}", "demand_index": "{:.3f}",
                                           "late_rate": "{:.1%}", "rain_share": "{:.1%}"}),
                     width="stretch")


# --------------------------------------------------------------------------------------
# Page 4 - Model comparison
# --------------------------------------------------------------------------------------
def page_comparison() -> None:
    header("Unit 5 — Model comparison on the held-out test period")
    need(U5 / "comparison_with_pbl1.csv", U5 / "feature_importance.csv", U5 / "test_predictions.csv",
         U1 / "test_predictions.csv", WEEKLY_CSV)
    comp = load_csv(str(U5 / "comparison_with_pbl1.csv"), index_col=0)
    imp = load_csv(str(U5 / "feature_importance.csv"), index_col=0)

    shown = comp.rename(index={GB_KEY: "Gradient Boosting"})
    shown = shown.loc[["Naive (last week)", "Linear Regression", "Ridge Regression", "Bayesian Regression",
                       "Gradient Boosting"], ["MAE", "RMSE", "R²"]]
    best = shown.MAE.idxmin()
    c1, c2, c3 = st.columns(3)
    c1.metric("Best model (lowest MAE)", best)
    c2.metric("Its MAE", f"{shown.loc[best, 'MAE']:.2f}", f"{shown.loc[best, 'MAE'] - shown.loc['Naive (last week)', 'MAE']:+.2f} vs naive",
              delta_color="inverse")
    c3.metric("Its R²", f"{shown.loc[best, 'R²']:.3f}")

    left, right = st.columns([1, 1.4])
    with left:
        st.markdown("##### Test metrics")
        st.dataframe(shown.style.format("{:.3f}").highlight_min(subset=["MAE", "RMSE"], color="rgba(27,175,122,.35)")
                     .highlight_max(subset=["R²"], color="rgba(27,175,122,.35)"), width="stretch")
        if "Gradient Boosting (original features)" in comp.index:
            st.caption(f"Gradient boosting on the original features (no growth-free target): "
                       f"MAE {comp.loc['Gradient Boosting (original features)', 'MAE']:.2f}.")
    with right:
        metric = st.radio("Metric", ["MAE", "RMSE", "R²"], horizontal=True)
        d = shown[metric].reset_index()
        d.columns = ["Model", metric]
        fig = px.bar(d, x="Model", y=metric, color="Model", color_discrete_map=MODEL_COLORS,
                     text=d[metric].round(3))
        fig.update_layout(showlegend=False, xaxis_title="Model",
                          yaxis_title=f"{metric} ({'lower' if metric != 'R²' else 'higher'} is better)")
        st.plotly_chart(style(fig, 340), width="stretch")

    st.markdown("##### Feature importance (gradient boosting)")
    kind = st.radio("Importance type", ["impurity importance", "permutation importance"], horizontal=True)
    fi = imp[kind].sort_values()
    fig = px.bar(x=fi.values, y=fi.index, orientation="h", text=fi.round(4))
    fig.update_traces(marker_color=BLUE)
    fig.update_layout(xaxis_title=kind.capitalize(), yaxis_title="Feature")
    st.plotly_chart(style(fig, 380), width="stretch")

    st.markdown("##### Festival weeks vs normal weeks (test period)")
    gb_pred = load_csv(str(U5 / "test_predictions.csv"), parse_dates=["week"])
    reg_pred = load_csv(str(U1 / "test_predictions.csv"), parse_dates=["week"])
    fest = load_weekly()[["store_id", "week", "festival_days"]]
    t = gb_pred.merge(reg_pred[["store_id", "week", "Naive (last week)", "Linear Regression"]],
                      on=["store_id", "week"]).merge(fest, on=["store_id", "week"])
    t["Week type"] = np.where(t.festival_days > 0, "Festival week", "Normal week")
    errs = pd.DataFrame({
        "Week type": t["Week type"],
        "Naive (last week)": (t.actual - t["Naive (last week)"]).abs(),
        "Linear Regression": (t.actual - t["Linear Regression"]).abs(),
        "Gradient Boosting": (t.actual - t.gb_growth_free).abs(),
    })
    mae = errs.groupby("Week type").mean().round(2)
    n = errs["Week type"].value_counts()
    long = mae.reset_index().melt(id_vars="Week type", var_name="Model", value_name="MAE")
    fig = px.bar(long, x="Week type", y="MAE", color="Model", barmode="group", color_discrete_map=MODEL_COLORS,
                 text="MAE")
    fig.update_layout(xaxis_title="Type of week", yaxis_title="MAE (orders per store-week)")
    a, b = st.columns([1.4, 1])
    a.plotly_chart(style(fig, 340), width="stretch")
    with b:
        st.dataframe(mae.style.format("{:.2f}"), width="stretch")
        st.caption(f"Store-weeks in the test set: {n.get('Normal week', 0):,} normal, {n.get('Festival week', 0):,} "
                   "festival. Festival weeks are rare, so their error estimate is less certain.")

    with st.expander("Charts from the notebook"):
        show_png(U5 / "comparison_mae.png", "Test MAE per model")
        show_png(U5 / "importance_and_learning_curve.png", "Feature importance and learning curve")


# --------------------------------------------------------------------------------------
# Page 5 - About
# --------------------------------------------------------------------------------------
def page_about() -> None:
    header("About this project")
    st.markdown(
        """
        ### AI Supply Chain Optimizer — Indian Quick-Commerce Food Demand
        **Course:** 21CSC305P Machine Learning — PBL 1 (Units 1–2), PBL 2 (Units 3–4), PBL 3 (Unit 5)

        **Dataset:** *India Quick Commerce Orders 2025–2026* (Kaggle, CC0) —
        [kaggle.com/datasets/abirambharathisa/india-quick-commerce-orders-2025-2026](https://www.kaggle.com/datasets/abirambharathisa/india-quick-commerce-orders-2025-2026).
        The food orders of 58 dark stores are turned into weekly store-level demand.

        | | |
        |---|---|
        | **Team** | *Team member names / register numbers — to be added* |
        | **Guide** | *Guide name and department — to be added* |
        | **Institution** | *Institution — to be added* |

        ### Supervised vs unsupervised methods
        **Supervised** methods learn from known answers (here, the number of orders actually received):
        *Linear, Ridge and Bayesian regression* (Units 1–2) and *gradient boosting* (Unit 5) are trained on past
        store-weeks to predict next week's orders, and are judged on weeks they have never seen.
        **Unsupervised** methods have no target and look for structure on their own: *K-Means* (Unit 3) groups the
        58 stores into three store types, *PCA* only projects the stores to two dimensions so that the groups can
        be drawn, and the *Hidden Markov Model* (Unit 4) discovers hidden daily demand regimes — Normal,
        Monsoon / disrupted and Festival peak — from the sequence of days.
        """
    )
    st.info("This app only loads the trained models and saved results; it never retrains anything.")


# --------------------------------------------------------------------------------------
# Navigation
# --------------------------------------------------------------------------------------
PAGES = {
    "Demand Forecast": page_forecast,
    "Store Types (K-Means)": page_clusters,
    "Demand Regimes (HMM)": page_regimes,
    "Model Comparison": page_comparison,
    "About": page_about,
}

ICONS = {"Demand Forecast": "📈", "Store Types (K-Means)": "🏪", "Demand Regimes (HMM)": "🌦️",
         "Model Comparison": "📊", "About": "ℹ️"}
st.sidebar.markdown('<div class="side-brand">🛵 Supply Chain Optimizer</div>'
                    '<div class="side-sub">Quick-commerce food demand</div>', unsafe_allow_html=True)
choice = st.sidebar.radio("Go to", list(PAGES), key="nav", format_func=lambda p: f"{ICONS[p]}  {p}",
                          label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.caption("Test MAE: gradient boosting 5.32 · naive 6.76 orders per store-week")
st.sidebar.caption("Loads saved models only — nothing is retrained.")

try:
    PAGES[choice]()
except Exception as exc:    # last line of defence: show a message instead of a stack trace
    st.error(f"Something went wrong on this page: {type(exc).__name__}: {exc}")

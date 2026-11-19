"""Builds the SRM 21CSC305P Machine Learning Project Report (.docx) from the template and the project results.

Run from the project folder:  python report/build_report.py
"""
import json, os, re, subprocess, sys, tempfile
import pandas as pd
import numpy as np
import nbformat

HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = HERE                                   # project folder = the nearest parent that contains dataset/
while not os.path.isdir(os.path.join(PROJ, "dataset")) and os.path.dirname(PROJ) != PROJ:
    PROJ = os.path.dirname(PROJ)
sys.path.insert(0, HERE)
from report_engine import Report, TEXT_W
from docx_helpers import J, C, L, R as RIGHT, fmt, add_rich, set_font, BLUE
from docx.shared import Inches, RGBColor

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "template.docx")
OUT_DOCX = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "ML_Project_Report_QuickCommerce.docx")
UNIT_OF = {"01_dataset": "Unit_1_2_PBL1_Data_and_Regression", "05_regression": "Unit_1_2_PBL1_Data_and_Regression",
           "02_clustering": "Unit_3_PBL2_Clustering_KMeans", "03_hmm": "Unit_4_PBL2_HMM", "04_comparison": "Unit_4_PBL2_HMM",
           "06_boosting": "Unit_5_PBL3_Gradient_Boosting"}
F = lambda sub, *p: os.path.join(PROJ, UNIT_OF[sub], "results", sub, *p)   # results live in the unit folders
FRONT = {
    "course": "21CSC305P – MACHINE LEARNING", "regulation": "2021 Regulation", "year_sem": "III Year/ V Semester",
    "academic_year": "2026 -2027",
    "team": [("Nikhil Saxena", "RA2411003011325"), ("Parnika Jain", "RA2411003011334"),
             ("Ayush Nischay", "RA2411003011347"), ("Neelansh Goyal", "RA2411003011345")],
    "guide": "Dr. Jeya R", "guide_desig": "Associate Professor",
    "hod": "Dr. G. Niranjana", "hod_desig": "Professor", "dept": "Computing Technologies",
    "branch": "COMPUTER SCIENCE AND ENGINEERING", "school": "SCHOOL OF COMPUTING", "month": "NOVEMBER 2026",
}
TITLE = ("AI Supply Chain Optimizer for Indian Quick-Commerce Food Demand: Design, Implementation, "
         "and Performance Evaluation of Machine Learning Techniques")

REFS = {
    "kaggle": 'A. Bharathi S. A., "India Quick Commerce Orders 2025–2026," Kaggle dataset, CC0 licence, 2026. [Online]. Available: https://www.kaggle.com/datasets/abirambharathisa/india-quick-commerce-orders-2025-2026',
    "sklearn": 'F. Pedregosa et al., "Scikit-learn: Machine learning in Python," J. Mach. Learn. Res., vol. 12, pp. 2825–2830, 2011.',
    "hmmlearn": 'hmmlearn developers, "hmmlearn: Hidden Markov models in Python," 2024. [Online]. Available: https://hmmlearn.readthedocs.io',
    "esl": 'T. Hastie, R. Tibshirani, and J. Friedman, The Elements of Statistical Learning, 2nd ed. New York, NY, USA: Springer, 2009.',
    "hoerl": 'A. E. Hoerl and R. W. Kennard, "Ridge regression: Biased estimation for nonorthogonal problems," Technometrics, vol. 12, no. 1, pp. 55–67, 1970.',
    "bishop": 'C. M. Bishop, Pattern Recognition and Machine Learning. New York, NY, USA: Springer, 2006.',
    "fpp": 'R. J. Hyndman and G. Athanasopoulos, Forecasting: Principles and Practice, 3rd ed. Melbourne, Australia: OTexts, 2021.',
    "lloyd": 'S. P. Lloyd, "Least squares quantization in PCM," IEEE Trans. Inf. Theory, vol. 28, no. 2, pp. 129–137, 1982.',
    "rousseeuw": 'P. J. Rousseeuw, "Silhouettes: A graphical aid to the interpretation and validation of cluster analysis," J. Comput. Appl. Math., vol. 20, pp. 53–65, 1987.',
    "rabiner": 'L. R. Rabiner, "A tutorial on hidden Markov models and selected applications in speech recognition," Proc. IEEE, vol. 77, no. 2, pp. 257–286, 1989.',
    "schwarz": 'G. Schwarz, "Estimating the dimension of a model," Ann. Stat., vol. 6, no. 2, pp. 461–464, 1978.',
    "geevers": 'K. Geevers, L. van Hezewijk, and M. R. K. Mes, "Multi-echelon inventory optimization using deep reinforcement learning," Cent. Eur. J. Oper. Res., vol. 32, no. 3, pp. 653–683, 2024, doi: 10.1007/s10100-023-00872-2.',
    "achkar": 'V. G. Achkar, B. B. Brunaud, R. Musa, and I. E. Grossmann, "Efficient computational strategies for a mathematical programming model for multi-echelon inventory optimization based on the guaranteed-service approach," Comput. Chem. Eng., vol. 182, Art. no. 108567, 2024, doi: 10.1016/j.compchemeng.2023.108567.',
    "claassen": 'G. D. H. Claassen, P. Kirst, A. T. T. Van, J. C. M. A. Snels, X. Guo, and P. van Beek, "Integrating time-temperature dependent deterioration in the economic order quantity model for perishable products in multi-echelon supply chains," Omega, vol. 125, Art. no. 103041, 2024, doi: 10.1016/j.omega.2024.103041.',
    "rizqi": 'Z. U. Rizqi and S.-Y. Chou, "Neuroevolution reinforcement learning for multi-echelon inventory optimization with delivery options and uncertain discount," Eng. Appl. Artif. Intell., vol. 134, Art. no. 108670, 2024, doi: 10.1016/j.engappai.2024.108670.',
    "friedman": 'J. H. Friedman, "Greedy function approximation: A gradient boosting machine," Ann. Stat., vol. 29, no. 5, pp. 1189–1232, 2001.',
}


APX_SETUP = '''import numpy as np, pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import (LinearRegression, RidgeCV,
                                  BayesianRidge)
from sklearn.cluster import KMeans, AgglomerativeClustering
from sklearn.mixture import GaussianMixture
from sklearn.decomposition import PCA
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import TimeSeriesSplit, GridSearchCV
from sklearn.inspection import permutation_importance
from sklearn.metrics import (mean_absolute_error,
    mean_squared_error, r2_score, silhouette_score,
    davies_bouldin_score, calinski_harabasz_score,
    adjusted_rand_score)
from hmmlearn.hmm import GaussianHMM
SEED = 42'''

APX_CH2 = '''# ---- 01_Create_Dataset: food filter and IPL-effect removal
FOOD = ["Fruits & Vegetables", "Snacks & Beverages",
        "Instant & Frozen Food", "Dairy, Bread & Eggs",
        "Atta, Rice & Dal", "Ice Cream & Desserts",
        "Sweets & Gifts"]
food = raw[raw.primary_category.isin(FOOD)].copy()  # 1,17,077
food["evening"] = food.order_hour.between(19, 22)
food["month"] = food.order_date.dt.month
is_ipl = food.event.eq("IPL match day")
sd = (food.groupby(["store_id", "order_date"])
          .agg(evening=("evening", "sum"),
               total=("order_id", "size"),
               event=("event", "first"),
               month=("month", "first")).reset_index())
sd["daytime"] = sd.total - sd.evening
reg = sd[sd.month.isin([3, 4, 5]) & sd.event.eq("Regular day")]
normal_ratio = (reg.groupby("store_id").evening.sum()
                / reg.groupby("store_id").daytime.sum())
ipl = (sd[sd.event.eq("IPL match day")]
       .groupby("store_id")[["evening", "daytime"]].sum())
ipl["expected"] = ipl.daytime * ipl.index.map(normal_ratio)
excess = (ipl.evening - ipl.expected) / ipl.evening
ipl["p_remove"] = excess.clip(0, 1)
rng = np.random.default_rng(SEED)
cand = food[is_ipl & food.evening]
p = cand.store_id.map(ipl.p_remove).values
drop = cand.order_id[rng.random(len(cand)) < p]
clean = food[~food.order_id.isin(drop)].copy()  # 1,13,390

# ---- 05_PBL1_Regression: store-week features (known in advance)
g = weekly.groupby("store_id").orders
weekly["lag1"] = g.shift(1)
weekly["lag2"] = g.shift(2)
weekly["lag4"] = g.shift(4)
weekly["roll4"] = g.transform(
    lambda s: s.shift(1).rolling(4).mean())
weekly["store_age_w"] = weekly.groupby("store_id").cumcount()
weekly["festival_days"] = weekly.week.map(fest_days).fillna(0)
weekly["monsoon"] = weekly.week.dt.month.isin(
    [6, 7, 8, 9, 10]) * 1
weekly = (pd.get_dummies(weekly, columns=["area_type"],
                         prefix="area", dtype=int)
          .drop(columns="area_Office / Tech Park").dropna())
FEATURES = ["lag1", "lag2", "lag4", "roll4", "store_age_w",
            "festival_days", "monsoon",
            "area_Premium Residential", "area_Residential",
            "area_Student / PG Hub"]

# time split; scaler fitted on the training weeks only
SPLIT = pd.Timestamp("2026-03-30")
train = weekly[weekly.week < SPLIT]
test = weekly[weekly.week >= SPLIT]
scaler = StandardScaler().fit(train[FEATURES])
Xs_train = scaler.transform(train[FEATURES])
Xs_test = scaler.transform(test[FEATURES])
y_train, y_test = train.orders, test.orders

# Unit 1: linear; Unit 2: ridge (time-series CV) and Bayesian
ols = LinearRegression().fit(Xs_train, y_train)
ridge = RidgeCV(alphas=np.logspace(-3, 3, 25),
                cv=TimeSeriesSplit(n_splits=5))
ridge.fit(Xs_train, y_train)
bayes = BayesianRidge().fit(Xs_train, y_train)
mu, sd = bayes.predict(Xs_test, return_std=True)
lo, hi = mu - 1.96 * sd, mu + 1.96 * sd     # 95% interval
coverage = ((y_test >= lo) & (y_test <= hi)).mean()

def metrics(y, p):
    return {"MAE": mean_absolute_error(y, p),
            "RMSE": mean_squared_error(y, p) ** 0.5,
            "R2": r2_score(y, p),
            "WAPE": np.abs(y - p).sum() / y.sum()}

results = pd.DataFrame({
    "Naive (last week)": metrics(y_test, test.lag1),
    "Linear Regression": metrics(y_test, ols.predict(Xs_test)),
    "Ridge Regression": metrics(y_test, ridge.predict(Xs_test)),
    "Bayesian Regression": metrics(y_test, mu)}).T'''

APX_U3 = '''# ---- 02_Clustering_KMeans
# signal-to-noise ratio: between-store variance divided by the
# variance expected from sampling noise alone
snr = pd.Series({f: stores[f].var() / noise_var(f)
                 for f in CANDIDATES})
CLUSTER_FEATURES = ["orders_per_day", "night_share",
    "morning_share", "share_fruit_veg", "share_snacks_bev",
    "share_dairy_bread_eggs", "share_atta_rice_dal",
    "share_icecream_desserts", "share_instant_frozen",
    "avg_basket_value"]
X = StandardScaler().fit_transform(stores[CLUSTER_FEATURES])

rows = []                                     # choice of k
for k in range(2, 9):
    km = KMeans(n_clusters=k, n_init=50, random_state=SEED)
    lab = km.fit_predict(X)
    rows.append({"k": k, "WCSS": km.inertia_,
                 "silhouette": silhouette_score(X, lab),
                 "DB": davies_bouldin_score(X, lab),
                 "CH": calinski_harabasz_score(X, lab)})

kmeans = KMeans(n_clusters=3, n_init=50, random_state=SEED)
labels = kmeans.fit_predict(X)
P = PCA(n_components=2, random_state=SEED).fit_transform(X)

# validation: other algorithms and stability over restarts
ward = AgglomerativeClustering(n_clusters=3, linkage="ward")
gmm = GaussianMixture(n_components=3, n_init=10,
                      random_state=SEED)
ari_ward = adjusted_rand_score(labels, ward.fit_predict(X))
ari_gmm = adjusted_rand_score(labels, gmm.fit(X).predict(X))
stability = np.mean([adjusted_rand_score(
    labels, KMeans(3, n_init=10, random_state=s).fit_predict(X))
    for s in range(20)])'''

APX_U4 = '''# ---- 01_Create_Dataset: demand index (growth, weekday removed)
daily["trend"] = daily.orders.rolling(
    29, center=True, min_periods=8).median()
log_ratio = np.log(daily.orders / daily.trend)
weekday_factor = (log_ratio[daily.event.eq("Regular day")]
                  .groupby(daily.dow).median())
daily["expected"] = daily.trend * np.exp(
    daily.dow.map(weekday_factor))
daily["demand_index"] = np.log(daily.orders / daily.expected)

# ---- 03_HMM: Gaussian HMM, Baum-Welch with 20 restarts
O = daily[["demand_index", "late_rate"]].values

def fit_hmm(obs, n_states, restarts=20):
    best = None
    for s in range(restarts):
        m = GaussianHMM(n_components=n_states,
                        covariance_type="diag", n_iter=500,
                        tol=1e-4, random_state=s).fit(obs)
        ll = m.score(obs)
        if best is None or ll > best[0]:
            best = (ll, m)
    return best

def n_params(n, d=2):      # pi + transitions + means + variances
    return (n - 1) + n * (n - 1) + 2 * n * d

bic = {k: n_params(k) * np.log(len(O)) - 2 * fit_hmm(O, k)[0]
       for k in range(2, 6)}
N_STATES = min(bic, key=bic.get)                # 3 states
_, hmm = fit_hmm(O, N_STATES)
states = hmm.predict(O)                         # Viterbi
post = hmm.predict_proba(O)                     # posteriors
expected_stay = 1 / (1 - np.diag(hmm.transmat_))
tomorrow = post[-1] @ hmm.transmat_             # next day

# out-of-sample check: train on 2025, score unseen 2026 days
tr = daily.index < "2026-01-01"
_, hmm_tr = fit_hmm(O[tr], N_STATES)
gmm_tr = GaussianMixture(N_STATES, covariance_type="diag",
                         n_init=20, random_state=SEED).fit(O[tr])
oos_hmm = hmm_tr.score(O[~tr]) / (~tr).sum()
oos_gmm = gmm_tr.score(O[~tr])

# ---- 04_Comparison: time-blind K-Means on the same days
Oz = StandardScaler().fit_transform(O)
km_days = KMeans(n_clusters=3, n_init=50,
                 random_state=SEED).fit_predict(Oz)'''

APX_U5 = '''# ---- 06_PBL3_Gradient_Boosting (same split/metrics as PBL 1)
for d in (train, test):            # relative, growth-free lags
    d["r1"] = d.lag1 / d.roll4
    d["r2"] = d.lag2 / d.roll4
    d["r4"] = d.lag4 / d.roll4
    d["woy"] = d.week.dt.isocalendar().week.astype(int)
REL = ["r1", "r2", "r4", "roll4", "store_age_w",
       "festival_days", "monsoon", "woy",
       "area_Premium Residential", "area_Residential",
       "area_Student / PG Hub"]
grid = {"n_estimators": [200, 400, 800],
        "learning_rate": [0.02, 0.05], "max_depth": [2, 3]}

def search():
    gbr = GradientBoostingRegressor(subsample=0.8,
                                    random_state=SEED)
    return GridSearchCV(gbr, grid, cv=TimeSeriesSplit(5),
                        scoring="neg_mean_absolute_error")

gb_raw = search().fit(train[FEATURES], train.orders)  # step 1
z_train = np.log(train.orders / train.roll4)  # growth-free
gb = search().fit(train[REL], z_train)                # step 2
pred = test.roll4 * np.exp(gb.predict(test[REL]))  # orders
gb_metrics = metrics(test.orders, pred)

best = gb.best_estimator_   # 200 trees, depth 2, rate 0.02
imp = pd.Series(best.feature_importances_, index=REL)
perm = permutation_importance(
    best, test[REL], np.log(test.orders / test.roll4),
    n_repeats=10, random_state=SEED,
    scoring="neg_mean_absolute_error")'''


def write(R, D):
    P = lambda t, **k: R.p(cite(t), **k)
    B = lambda items: R.bullets([cite(i) for i in items])
    m, comp = D["metrics"], D["comp"]
    lin, nav = m.loc["Linear Regression"], m.loc["Naive (last week)"]
    gbg = comp.loc["Gradient Boosting (growth-free target)"]; gbr = comp.loc["Gradient Boosting (original features)"]
    km, h, rg = D["km"], D["hmm"], D["reg"]

    # ================================================================== PRELIMINARY PAGES
    R.prelim_heading("ABSTRACT", page_break=False)
    abstract = (
        "Quick-commerce platforms in India deliver groceries within minutes from small urban dark stores, so every "
        "store must be stocked and staffed for food demand that grows quickly, differs between neighbourhoods and "
        "changes with weather and festivals. The integrated machine learning system developed here, the AI Supply "
        "Chain Optimizer, supports these decisions. The objectives were to forecast the weekly orders of each "
        "store, to group stores with similar demand, to detect hidden demand regimes over time and to test whether an "
        "ensemble model improves the forecast. The original dataset is the public Kaggle dataset India Quick Commerce "
        "Orders 2025–2026, with 1,95,148 orders from 58 dark stores in 12 cities; after cleaning, retaining seven food "
        "categories and removing the extra evening orders caused by IPL matches, 1,13,390 orders remained; Unit 4 used "
        "a 638-day sequence derived from the same data. In PBL 1, linear, ridge and Bayesian linear regression were "
        "trained on 2,531 store-weeks and tested on the next 26 weeks, reducing the mean absolute error from "
        f"{nav.MAE:.2f} orders for a naive forecast to {lin.MAE:.2f} (R² = {lin['R²']:.2f}). In PBL 2, K-Means grouped "
        "the stores into three types (silhouette 0.42, stability 0.85), and a three-state Gaussian hidden Markov model "
        "chosen by BIC separated normal, monsoon-disrupted and festival-peak regimes and detected 82% of festival days. "
        f"In PBL 3, gradient boosting with a growth-free target gave the lowest error (MAE {gbg.MAE:.2f}, "
        f"R² = {gbg['R²']:.3f}). Together, the models show where demand differs, when conditions change and how much "
        "each store should be replenished.")
    assert 150 <= len(abstract.split()) <= 250, len(abstract.split())
    R.p(abstract)

    R.prelim_heading("TABLE OF CONTENTS"); R.placeholder("toc")
    R.prelim_heading("LIST OF FIGURES"); R.placeholder("lof")
    R.prelim_heading("LIST OF TABLES"); R.placeholder("lot")
    R.prelim_heading("ABBREVIATIONS")
    abbr = [("ARI", "Adjusted Rand Index"), ("BIC", "Bayesian Information Criterion"),
            ("CART", "Classification and Regression Trees"), ("CH", "Calinski–Harabasz index"),
            ("CV", "Cross-Validation"), ("DB", "Davies–Bouldin index"), ("EDA", "Exploratory Data Analysis"),
            ("EM", "Expectation–Maximisation"), ("EOQ", "Economic Order Quantity"), ("GB", "Gradient Boosting"), ("GMM", "Gaussian Mixture Model"), ("GSM", "Guaranteed-Service Model"),
            ("HMM", "Hidden Markov Model"), ("IPL", "Indian Premier League (cricket)"), ("MAE", "Mean Absolute Error"),
            ("MIQCP", "Mixed-Integer Quadratically Constrained Program"), ("NERL", "Neuroevolution Reinforcement Learning"),
            ("OLS", "Ordinary Least Squares"), ("PBL", "Project-Based Learning"), ("PCA", "Principal Component Analysis"),
            ("PG", "Paying Guest (student housing)"), ("PPO", "Proximal Policy Optimization"), ("R²", "Coefficient of determination"),
            ("RL", "Reinforcement Learning"), ("RMSE", "Root Mean Squared Error"), ("SNR", "Signal-to-Noise Ratio"),
            ("WAPE", "Weighted Absolute Percentage Error"), ("WCSS", "Within-Cluster Sum of Squares")]
    R.table("", ["ABBREVIATION", "EXPANSION"], abbr, [1.4, 3.8], align=[L, L], numbered=False, bold_first=True)

    R.new_main_section()

    # ================================================================== CHAPTER 1
    R.chapter("Introduction", first=True)
    R.h2("1.1 Background, Motivation and Scope")
    P("Quick commerce (q-commerce) platforms such as Blinkit, Zepto and Swiggy Instamart deliver groceries in about ten "
      "minutes. Each city is served by small *dark stores* — closed warehouses inside residential areas, student hubs "
      "and office districts — from which riders deliver within two to three kilometres. This moves the hardest supply-"
      "chain decisions to the individual store: shelf space is small, fresh food spoils within days, and demand changes "
      "with the neighbourhood, the hour, the weather and the festival calendar. Under-stocking causes cancellations, "
      "over-stocking wastes food, and a wrong number of riders causes late deliveries or idle cost.")
    P("The stakeholders are the central replenishment team (weekly order forecast per store), dark-store managers "
      "(product mix and staffing), fleet managers (rider shifts and promised times) and customers. **Scope:** one public "
      "order-level dataset [@kaggle], limited to **food** categories (fruits and vegetables, dairy, staples, snacks and "
      "beverages, instant and frozen food, desserts and sweets), with the artificial evening demand of IPL cricket "
      "matches removed. The work covers data preparation, EDA, regression, clustering, sequential modelling and "
      "ensemble learning in Python with scikit-learn and hmmlearn [@sklearn, @hmmlearn]; real-time deployment and "
      "product-level (SKU) planning are outside the scope.")
    R.h2("1.2 Integrated Problem Statement")
    P("**Overall problem:** *How can a quick-commerce operator use its order history to plan food replenishment and "
      "rider capacity for every dark store?* It is divided into linked sub-problems on the same original dataset:")
    B(["**Units 1–2 (PBL 1):** forecast each store's food orders for the next week from information available at the "
       "end of the current week (supervised regression).",
       "**Unit 3 (PBL 2):** group the 58 dark stores by demand profile so that a group can share one stocking and "
       "staffing plan (clustering).",
       "**Unit 4 (PBL 2):** using a 638-day sequence derived from the same orders (no separate dataset), discover the "
       "hidden demand regimes of the network, how long they last and which is likely tomorrow (HMM).",
       "**Unit 5 (PBL 3):** re-solve the PBL 1 task with gradient boosting on identical features, test weeks and metrics."])
    R.h2("1.3 Project Objectives")
    B(["**Objective 1 (Unit 1):** build a clean, leakage-free store-week dataset and a linear regression that beats the "
       "naive forecast — *achieved: MAE 5.35 vs 6.76.*",
       "**Objective 2 (Unit 2):** apply ridge (time-series CV penalty) and Bayesian regression and obtain prediction "
       "intervals — *achieved: equal accuracy, intervals covering 90.6% of test weeks.*",
       "**Objective 3 (Unit 3):** select reliable store features, choose k and interpret the K-Means clusters — "
       "*achieved: three stable store types, silhouette 0.42.*",
       "**Objective 4 (Unit 4):** train a Gaussian HMM with Baum–Welch, choose the states by BIC, decode with Viterbi "
       "and validate on unseen days — *achieved: three regimes, 82% festival recall.*",
       "**Objective 5 (Unit 5):** tune gradient boosting with time-series CV and compare it with PBL 1 — *achieved: "
       "lowest test MAE (5.32).*"])
    R.h2("1.4 Unit-wise Topic and Dataset Allocation")
    P(f"{R.next_tab()} lists the topic and dataset of every unit. Unit 4 uses a daily sequence derived from the original "
      "dataset, so that its regimes can be combined with the store clusters of Unit 3 (Chapter 6).", keep_next=True)
    R.table("Unit-wise topic and dataset allocation",
            ["PBL", "Unit", "Allotted topic / algorithm", "Dataset"],
            [["PBL 1", "1", "Data preprocessing, EDA and Linear Regression (OLS)", "Original"],
             ["PBL 1", "2", "Ridge Regression and Bayesian Linear Regression", "Original"],
             ["PBL 2", "3", "Clustering: K-Means (PCA for visualisation)", "Original"],
             ["PBL 2", "4", "Hidden Markov Model: Baum–Welch, Viterbi", "Original (daily sequence)"],
             ["PBL 3", "5", "Ensemble learning: Gradient Boosting (CART trees)", "Original"]],
            [0.65, 0.5, 3.0, 1.61], align=[C, C, L, L])

    # ================================================================== CHAPTER 2 – LITERATURE SURVEY
    R.chapter("Literature Survey")
    R.h2("2.1 Background and Problem Context")
    P("A quick-commerce network is a **multi-echelon** supply chain: a central supply point feeds regional hubs, the hubs "
      "feed distribution nodes (here, dark stores), and the nodes serve customer demand. Every inventory decision in such "
      "a network has to balance three groups of factors:")
    B(["**Uncertainty** — demand forecasts are wrong at every node, upstream deliveries are not fully reliable and "
       "replenishment lead times vary.",
       "**Inventory and capacity** — too much stock locks up capital, too little causes lost sales and backorders, and "
       "every node has only a finite amount of storage space.",
       "**Product and cost** — fresh products lose quality during storage and expire after a fixed shelf life, while "
       "holding stock costs money for storage, cooling and handling."])
    P("Traditional rule-based systems set the stock of each node separately and treat these factors one at a time. The "
      "survey therefore looked for methods that consider demand, inventory, capacity, safety stock, product "
      "characteristics and the network structure together. Four recent studies were selected, each covering one part of "
      "this problem: learned ordering policies, safety-stock placement, perishable products and decisions under "
      "uncertainty. They are reviewed below under the same headings — problem, method, evaluation and limitations.")

    R.h2("2.2 Review of Related Work")
    R.h3("2.2.1 Deep reinforcement learning for multi-echelon inventory control")
    P("Geevers et al. [@geevers] ask how well deep reinforcement learning (RL) can control inventory in a multi-echelon "
      "network, where an order placed at one stock point changes what every other stock point needs. Instead of applying "
      "a separate ordering rule at each node, a single agent manages the whole network. The agent observes the state of "
      "every echelon — stock on hand, backorders and goods in transit — and outputs a continuous order quantity for each "
      "stock point. A simulator then lets demand arrive and shipments move forward by one period, and the period's "
      "holding and backorder cost is returned as a negative reward. The policy is trained with **Proximal Policy "
      "Optimization (PPO)**, which repeatedly shifts the policy towards actions that earned a higher reward.")
    P("The method was tested on three structures: a linear and a divergent network taken from the literature, and a "
      "general network based on a real manufacturer. The PPO policy beat the published benchmark in all three cases, with "
      "average cost reductions of 16.4%, 11.3% and 6.6% respectively [@geevers]. The gain is largest in the simplest network "
      "and smallest in the realistic one. Perishability, product quality and warehouse capacity are not modelled, and a "
      "learned policy is harder to explain to planners than a fixed rule.")

    R.h3("2.2.2 Safety-stock placement with the guaranteed-service model")
    P("Achkar et al. [@achkar] address where to hold safety stock in a chain with many locations and products, so that "
      "service targets (fill rate or cycle service level) are met at minimum cost. In the guaranteed-service model (GSM) "
      "every stage promises a service time to its customers; the stage must then cover demand variability only over its "
      "*net replenishment time* — its inbound service time plus its own processing time, minus the service time it "
      "promises downstream. This formulation is nonlinear and nonconvex and therefore slow to solve. The authors "
      "reformulate it as a **mixed-integer quadratically constrained program (MIQCP)**, replace the nonlinear fill-rate "
      "curve with a piecewise-linear approximation and represent stochastic lead times with a discrete function.")
    P("On several instances of a real pharmaceutical supply chain with more than 7,300 product–location combinations, "
      "the strategy found optimal safety-stock plans within a few seconds and estimated the fill rate more accurately "
      "than the original formulation [@achkar]. The approach is a periodic, tactical planning tool rather than a "
      "real-time allocation method, and it does not model perishability or storage temperature.")

    R.h3("2.2.3 Time–temperature deterioration of perishable products")
    P("Claassen et al. [@claassen] observe that the classical economic order quantity (EOQ) model for deteriorating "
      "items assumes one fixed rate of decay, although the quality of fresh products depends on how long and at what "
      "temperature they are kept at each stage. In their model each echelon of a cold chain (transport, cold storage, "
      "distribution, retail) is described by its duration and temperature; a product-specific function converts these "
      "into quality loss, the losses are accumulated along the route, and the resulting deterioration enters the "
      "ordering, holding and shortage cost balance of the EOQ.")
    P("The methodology was demonstrated on a real cold-chain case from floriculture, comparing two commonly used "
      "multi-echelon chains [@claassen]. It shows that chain design and storage conditions change the quality that "
      "reaches the customer and therefore the best order quantity and cycle. It builds on EOQ assumptions, so demand "
      "is treated as known, and safety stock and demand uncertainty are not included.")

    R.h3("2.2.4 Neuroevolution reinforcement learning under uncertainty")
    P("Rizqi and Chou [@rizqi] study a centralised ordering decision that must cope with uncertain demand, uncertain "
      "supply and unexpected supplier discounts, while suppliers offer delivery options that trade cost against lead "
      "time. The ordering policy is a neural network whose weights are not trained by gradient descent but **evolved**: "
      "a population of policies is run in a system-dynamics simulation of the chain, each policy's total system cost "
      "is its fitness, and differential evolution (DE), a memetic algorithm (MA) or an evolution strategy (ES) "
      "recombines the best policies into the next generation. Each policy chooses both the order quantity and the "
      "delivery option in every period.")
    P("In experiments on serial and divergent networks, this neuroevolution RL (NERL) framework outperformed an "
      "EA-optimised continuous-review (s, Q) policy with every evolutionary algorithm tested [@rizqi]. The evaluation "
      "is carried out in simulation, training is computationally expensive, and perishability, storage capacity and "
      "product quality are not modelled.")

    R.h2("2.3 Comparative Analysis")
    P(f"{R.next_tab()} compares the four studies on a common set of criteria.", keep_next=True)
    R.table("Comparison of the reviewed studies",
            ["Study", "Method", "Evaluation", "Key result", "Not covered"],
            [[cite("Geevers et al. [@geevers]: ordering policies"), "Deep RL (PPO) with continuous actions",
              "Linear, divergent and real general network", "16.4%, 11.3% and 6.6% lower cost than benchmark",
              "Perishability, capacity, explainability"],
             [cite("Achkar et al. [@achkar]: safety stock"), "Guaranteed-service model as MIQCP",
              "Pharmaceutical chain, 7,300+ product–locations", "Optimal plans in a few seconds",
              "Real-time decisions, perishability"],
             [cite("Claassen et al. [@claassen]: perishables"), "EOQ with time–temperature decay",
              "Two floriculture cold chains", "Quality-aware order quantity",
              "Demand uncertainty, safety stock"],
             [cite("Rizqi and Chou [@rizqi]: uncertainty"), "Neuroevolution RL (DE, MA, ES)",
              "Serial and divergent networks, simulated", "Beats an EA-optimised (s, Q) policy",
              "Perishability, capacity, real data"]],
            [1.15, 1.2, 1.2, 1.15, 1.06], align=[L, L, L, L, L])
    P("The studies fall into two families. Optimisation models [@achkar, @claassen] give exact, explainable answers but "
      "rely on assumed demand distributions and fixed parameters. Learning-based policies [@geevers, @rizqi] adapt to the "
      "network state and to uncertainty and outperform tuned classical rules, but they need a simulator, are expensive "
      "to train and are harder to explain. Only one study treats product quality, and only one treats service levels "
      "explicitly, so no single method covers the whole problem described in Section 2.1.")

    R.h2("2.4 Research Gaps")
    B(["**Demand is assumed, not learned.** All four studies take demand as a given distribution or simulate it. None "
       "estimates it from real transaction data, although the quality of every ordering and safety-stock decision "
       "depends on how well demand at each node is forecast.",
       "**Nodes are treated alike.** Stock points differ only in their position in the network; differences in demand "
       "profile between locations (for example residential versus office areas) are not used to group them.",
       "**Regime changes are treated as noise.** Festival peaks and seasonal disruptions are absorbed into random "
       "demand instead of being recognised as recurring demand states that can be anticipated.",
       "**Perishability is handled separately.** Product quality appears only in the EOQ study; the RL and safety-"
       "stock methods ignore shelf life."])

    R.h2("2.5 Positioning of This Project")
    P("The survey led to the design of a supply-chain optimizer with six modules: product categorisation, demand "
      "forecasting, a multi-echelon network model, an inventory state monitor (stock, capacity and quality), safety "
      "stock, and an optimisation engine that decides allocation and replenishment. This machine-learning project builds "
      f"the data-driven modules that the reviewed optimisation methods take as given, as summarised in {R.next_tab()}.")
    P("In short, the literature shows *how* to decide once demand is known; this project addresses *what the demand "
      "will be, where it differs and when it changes*, using the original quick-commerce order data. Its forecasts, "
      "store types and regimes are the inputs that an optimisation engine of the kind reviewed above would need.")
    R.table("Optimizer modules, supporting literature and coverage in this project",
            ["Optimizer module", "Basis in the literature", "Covered in this project"],
            [["Demand forecasting per node", cite("Input assumed known in [@geevers, @rizqi, @claassen]"),
              "Units 1, 2 and 5: weekly order forecast for every dark store (Chapters 3 and 5)"],
             ["Safety stock and service levels", cite("Guaranteed-service model [@achkar]"),
              "Unit 2: Bayesian 95% prediction intervals give the demand spread a safety-stock rule needs (Chapter 3)"],
             ["Node categorisation and class-specific rules", cite("Product classes and handling rules [@claassen]"),
              "Unit 3: K-Means groups the 58 stores into three store types (Chapter 4)"],
             ["Uncertainty and demand regimes", cite("Uncertain demand and supply [@rizqi]"),
              "Unit 4: HMM detects Normal, Monsoon and Festival-peak regimes (Chapter 4)"],
             ["Optimisation engine (allocation and replenishment)", cite("Learned policies and exact models [@geevers, @achkar, @rizqi]"),
              "Not implemented; future scope (Chapter 7)"],
             ["Perishability and quality", cite("Time–temperature deterioration [@claassen]"),
              "Scope limited to food categories; shelf life is not in the dataset (future scope)"]],
            [1.55, 1.75, 2.46], align=[L, L, L])
    R.h2("2.6 Chapter Conclusion")
    P("The reviewed studies show that multi-echelon inventory decisions can be improved substantially by learned "
      "ordering policies [@geevers, @rizqi], by exact safety-stock optimisation [@achkar] and by modelling product "
      "deterioration [@claassen]. All of them, however, start from demand that is already known or simulated. This "
      "project therefore concentrates on the demand side, using real quick-commerce orders: Chapter 3 forecasts weekly "
      "store demand with linear, ridge and Bayesian regression; Chapter 4 groups the stores with K-Means and detects "
      "demand regimes with a hidden Markov model; Chapter 5 improves the forecast with gradient boosting. Chapter 7 "
      "returns to the optimisation engine as future scope.")

    # ================================================================== CHAPTER 3
    R.chapter("Data Analysis and Fundamental Machine Learning Techniques – (Units 1 and 2)")
    R.h2("3.1 Overview of Allotted Topics")
    P("**Unit 1 — preprocessing, EDA and linear regression.** Preprocessing turns the raw orders into a clean, "
      "leakage-free modelling table and EDA shows which variables carry signal. Linear regression then models the "
      "weekly food orders of a store as a weighted sum of features [@esl]:")
    R.eq(r"\hat{y} = w_0 + \sum_{j=1}^{p} w_j x_j = w_0 + \mathbf{w}^{\top}\mathbf{x}")
    P("It suits the problem because next week's orders are almost linearly related to recent orders (correlation 0.88 "
      "with the four-week mean) and its coefficients are directly readable by planners. It assumes a linear relation, "
      "independent errors with constant variance and no perfect multicollinearity. **Unit 2 — ridge and Bayesian "
      "regression.** The lag features are strongly correlated, which can make least-squares weights unstable; ridge "
      "regression shrinks them with an L2 penalty [@hoerl]. Bayesian regression places a Gaussian prior on the weights "
      "and returns a predictive distribution whose spread can be turned into safety stock [@bishop]. A naive forecast "
      "(next week = this week) is the reference every model must beat [@fpp].")

    R.h2("3.2 Dataset Description")
    dd = D["dict"]
    P(f"{R.next_tab()} summarises the original dataset: one row per order with its time, store, basket, payment and "
      "delivery outcome; Table 3.2 groups its 39 attributes.", keep_next=True)
    R.table("Summary of the original dataset",
            ["Property", "Value"],
            [["Name / source", cite("India Quick Commerce Orders 2025–2026, Kaggle [@kaggle] (CC0)")],
             ["Version", "Release covering 01-01-2025 to 30-09-2026 (638 days)"],
             ["Size", "1,95,148 rows × 39 columns; 12 cities, 58 dark stores, 46,804 customers"],
             ["Data types", f"{(dd.dtype.isin(['int64','float64'])).sum()} numeric, {(dd.dtype=='bool').sum()} Boolean, "
                            f"{(dd.dtype=='str').sum()} categorical / text, {dd.dtype.str.startswith('datetime').sum()} date-time"],
             ["Target", "None in the raw data; derived target = food orders of a store in a week"],
             ["Missing values", "5 columns, all structural (Table 3.3)"]],
            [1.4, 4.36], align=[L, L], bold_first=True)
    rows = [[g, str(len(sub)), ", ".join(sub.column)] for g, sub in dd.groupby("group", sort=False)]
    R.table("Attribute groups of the dataset", ["Group", "No.", "Attributes"], rows, [1.3, 0.45, 4.01], align=[L, C, L])
    R.table("Missing values in the raw dataset (all structural)",
            ["Column", "Missing", "Reason"],
            [["actual_delivery_minutes, is_late", "3,143 (1.6%)", "cancelled orders were never delivered"],
             ["cancel_reason", "1,92,005 (98.4%)", "filled only for cancelled orders"],
             ["complaint_type", "1,85,367 (95.0%)", "filled only when there was a complaint"],
             ["customer_rating", "1,22,204 (62.6%)", "customer did not rate the order"]],
            [2.25, 1.35, 2.16], align=[L, C, L])
    P("**Limitations.** The data is simulated to resemble real quick-commerce operations, so the numbers illustrate the "
      "methods rather than a real company; each order has one *main* category; the IPL effect is embedded and must be "
      "removed statistically; 21 months give only two of each season; and there is no inventory or price data.")

    R.h2("3.3 Data Preprocessing")
    P("**Quality checks.** There are no duplicate orders or negative amounts, and the payment formula (item total − "
      "discount − refund + fees + tip) holds for 98.4% of orders — exactly the delivered ones (cancelled orders paid "
      "nothing). Missing values were kept as missing, never imputed: delivery time and lateness are computed over "
      "delivered orders only. **Food filter.** Orders whose main category is one of seven food categories were kept "
      "(1,17,077 of 1,95,148, 60.0%); 78,071 orders of personal care, household, baby care, electronics, pharmacy, pet "
      "care and flowers were removed.")
    P("**Removal of the IPL effect.** On the 128 IPL match days (22-03-2025 to 31-05-2026) daytime demand is unchanged "
      "(2.59 vs 2.55 orders per store-day on regular March–May days) but evening orders (7–11 pm) rise from 1.10 to "
      "1.81. For each store *s* the normal evening-to-daytime ratio *r*_{s}, the expected IPL-evening orders and the "
      "removal probability are")
    R.eq(r"r_s = \frac{E_s^{\text{reg}}}{D_s^{\text{reg}}}, \qquad \hat{E}_s = r_s\, D_s^{\text{IPL}}, \qquad p_s = \max\!\left(0,\; \frac{E_s^{\text{IPL}} - \hat{E}_s}{E_s^{\text{IPL}}}\right)")
    P(f"where *E* and *D* are evening and daytime orders. Each IPL-evening order was dropped with probability *p*_{{s}} "
      f"(seed 42): 3,687 orders were removed, leaving **1,13,390 clean food orders**. {R.next_fig()} shows that the evening "
      "share on IPL days falls from 41.1% to 30.4%, the level of regular days (30.2%).")
    R.fig(F("01_dataset", "ipl_removal.png"), "Share of orders placed 7–11 pm in March–May before and after IPL removal", 4.4)
    P("**Store-week dataset and features.** Orders were counted per store and week (weeks start on Monday); each "
      "store's partly open first week and the incomplete last week were dropped. All features are known before the "
      f"week starts, so no future information leaks into a forecast ({R.next_tab()}). Location type was one-hot encoded "
      "with Office / Tech Park as the reference category; city was not used.", keep_next=True)
    R.table("Features of the store-week dataset (4,039 rows, 58 stores, 03-02-2025 to 21-09-2026)",
            ["Feature", "Description", "Type"],
            [["`orders` (target)", "food orders of the store in week *t*", "count"],
             ["`lag1`, `lag2`, `lag4`", "orders 1, 2 and 4 weeks earlier", "count"],
             ["`roll4`", "mean orders of the previous 4 weeks", "continuous"],
             ["`store_age_w`", "weeks since the store opened (growth)", "integer"],
             ["`festival_days`", "festival days in the week (0–3), from the calendar", "integer"],
             ["`monsoon`", "1 for weeks in June–October", "binary"],
             ["`area_*` (3)", "location: Premium Residential, Residential, Student / PG Hub", "binary"]],
            [1.6, 3.26, 0.9], align=[L, L, C])
    P("**Train/test split and scaling.** The data is a time series, so it was split by time: training weeks 03-02-2025 "
      "to 23-03-2026 (2,531 store-weeks) and test weeks 30-03-2026 to 21-09-2026 (26 weeks, 1,508 store-weeks). The "
      "z-score scaler below was fitted on the training weeks only, so no test information leaks into training:")
    R.eq(r"z_{ij} = \frac{x_{ij} - \mu_j}{\sigma_j}")

    R.h2("3.4 Exploratory Data Analysis (EDA)")
    P(f"**Interpretation of {R.next_fig()}.** Daily orders grow from about 76 a day in the first month to about 500 in the last month, with "
      "sharp festival spikes, so every model must handle growth. The kept food categories include the three largest "
      "(staples, snacks and beverages, instant and frozen food). Orders peak at 8–10 pm and stay high late at night, "
      "while 2–5 am is nearly empty — a daily rhythm that differs by store and is used in Unit 3.")
    R.fig(F("01_dataset", "eda_overview.png"), "Raw data overview: daily orders, orders by main category and by hour of day", 5.3)
    P(f"**Interpretation of {R.next_fig()}.** The target is right-skewed (mean 26.9, median 25, maximum 92). `roll4` has the strongest "
      "correlation with next week's orders (0.88), followed by the lags (about 0.81), store age (0.32), monsoon (0.21) and "
      "festival days (0.10); festival weeks average 29.5 orders against 26.6. All location types grow at a similar pace, "
      "student hubs being the busiest and most volatile. Recent demand will therefore dominate any forecast, while "
      "calendar features add smaller corrections, and the inter-correlated lags motivate ridge regularisation.")
    R.fig(F("05_regression", "eda_weekly.png"), "Weekly orders per store: distribution, correlation with the features and trend by location", 5.3)

    R.h2("3.5 Proposed System Architecture / Workflow")
    P(f"{R.next_fig()} shows the workflow. " "The raw data is cleaned, filtered to food and corrected for IPL (stages 1–4), turned into the store-week table "
      "(5) and split by time with training-only scaling (6); four models are trained (7), evaluated on 26 unseen weeks "
      "(8) and deliver a next-week forecast with uncertainty for every store (9).")
    R.fig(os.path.join(HERE, "fig_workflow_ch2.png"), "Workflow of PBL 1 — weekly store demand forecasting")

    R.h2("3.6 Algorithm Theory and Mathematical Formulation")
    P("**Notation:** *n* = 2,531 training store-weeks, *p* = 10 standardised features, **X** the *n* × *p* feature matrix, "
      "**y** the weekly orders and **w** the weights. The **naive baseline** is ŷ_{t} = *y*_{t−1}. **Linear regression "
      "(OLS)** minimises the residual sum of squares; **ridge regression** adds an L2 penalty α‖**w**‖² [@hoerl]:")
    R.eq(r"\hat{\mathbf{w}}_{\text{OLS}} = (\mathbf{X}^{\top}\mathbf{X})^{-1}\mathbf{X}^{\top}\mathbf{y}, \qquad \hat{\mathbf{w}}_{\text{ridge}} = \arg\min_{\mathbf{w}} \lVert \mathbf{y} - \mathbf{X}\mathbf{w} \rVert^2 + \alpha \lVert \mathbf{w} \rVert^2 = (\mathbf{X}^{\top}\mathbf{X} + \alpha \mathbf{I})^{-1}\mathbf{X}^{\top}\mathbf{y}")
    P("α was chosen from 25 values between 10^{−3} and 10^{3} by **5-fold time-series cross-validation**, in which each "
      "fold trains on earlier weeks and validates on the following ones. **Bayesian linear regression** [@bishop] uses a "
      "Gaussian prior of precision λ on the weights and Gaussian noise of precision α:")
    R.eq(r"p(\mathbf{w}) = \mathcal{N}(\mathbf{0}, \lambda^{-1}\mathbf{I}), \quad p(\mathbf{y} \mid \mathbf{X}, \mathbf{w}) = \mathcal{N}(\mathbf{X}\mathbf{w}, \alpha^{-1}\mathbf{I}), \quad \mathbf{S} = (\lambda \mathbf{I} + \alpha \mathbf{X}^{\top}\mathbf{X})^{-1}, \quad \mathbf{m} = \alpha \mathbf{S} \mathbf{X}^{\top}\mathbf{y}")
    P("The posterior of the weights is N(**m**, **S**) and the predictive distribution of a new week is "
      "N(**m**^{⊤}**x**, α^{−1} + **x**^{⊤}**Sx**); α and λ are estimated by maximising the marginal likelihood, and "
      "mean ± 1.96 standard deviations gives a 95% prediction interval. **Metrics** (*N* test store-weeks):")
    R.eq(r"\text{MAE} = \frac{1}{N}\sum_{i} \left| y_i - \hat{y}_i \right|, \quad \text{RMSE} = \sqrt{\frac{1}{N}\sum_{i} (y_i - \hat{y}_i)^2}, \quad R^2 = 1 - \frac{\sum_{i}(y_i - \hat{y}_i)^2}{\sum_{i}(y_i - \bar{y})^2}, \quad \text{WAPE} = \frac{\sum_{i}\left| y_i - \hat{y}_i \right|}{\sum_{i} y_i}")
    P("MAE (orders per store-week) is the main metric because it is directly usable for replenishment; RMSE penalises "
      "large misses, R² is the variance explained and WAPE the total error as a share of demand. These methods fit "
      "because the relation is nearly linear, planners need explainable weights, ridge guards against correlated lags "
      "and the Bayesian model adds the uncertainty needed for safety stock.")

    R.h2("3.7 Results and Performance Evaluation")
    P(f"All models were evaluated on the same 1,508 test store-weeks ({R.next_tab()}).", keep_next=True)
    R.table("Test performance of the PBL 1 models (26 unseen weeks)",
            ["Model", "MAE ↓", "RMSE ↓", "R² ↑", "WAPE ↓"],
            [[i, f"{r.MAE:.3f}", f"{r.RMSE:.3f}", f"{r['R²']:.3f}", f"{r.WAPE:.1%}"] for i, r in m.iterrows()],
            [2.3, 0.85, 0.85, 0.85, 0.91])
    P(f"All three regression models clearly beat the naive forecast: MAE falls from {nav.MAE:.2f} to {lin.MAE:.2f} orders "
      f"(−{1 - lin.MAE / nav.MAE:.0%}), R² rises from {nav['R²']:.2f} to {lin['R²']:.2f} and WAPE falls from "
      f"{nav.WAPE:.1%} to {lin.WAPE:.1%}. The three models are practically identical: cross-validation chose the smallest "
      f"ridge penalty tried (α = {rg['ridge_alpha']:g}), and the Bayesian evidence gave a noise precision of 0.0295 "
      "(noise std ≈ 5.8 orders) and a weak prior (λ = 0.1508). With 2,531 rows and ten features there is little "
      f"variance for regularisation to remove. The Bayesian 95% interval covers {rg['bayes_coverage_95']:.1%} of test weeks.")
    P(f"**Interpretation of {R.next_fig()}.** " "`roll4` dominates (+8.4 orders per standard deviation): next week mostly continues the recent "
      "level. Festival days add +1.87 per standard deviation, about **2.6 extra orders per festival day**. Store age is "
      "slightly negative because the lags already contain growth, and location effects are small once recent demand is "
      "known. The ridge path is flat until α ≈ 10, so the least-squares solution is already stable.")
    R.fig(F("05_regression", "coefficients.png"), "Standardised model coefficients and the ridge coefficient path (dashed line = chosen α)")
    P(f"**Interpretation of {R.next_fig()}.** " "Predictions follow the diagonal with a spread that widens for busy stores. Most weekly values "
      "of store DS001 fall inside the Bayesian interval. At network level the naive forecast is always one week late, "
      "most visibly at the August 2026 Raksha Bandhan peak, while linear regression reacts to the festival calendar but "
      "stays slightly below actual demand during rapid growth (mean residual +1.67 orders, std 6.77).")
    R.fig(F("05_regression", "forecast_results.png"), "Test-period forecasts: actual vs predicted, Bayesian 95% interval for store DS001 and network total per week", 5.0)

    R.h2("3.8 Discussion and Findings")
    P("Leakage-free features explain 78% of the variance on unseen weeks. Regularisation did not help because data is "
      "plentiful relative to the features. Coverage is 90.6% (not 95%) because count noise grows with the level; a count "
      "model would calibrate it. Challenges were the IPL effect, new stores and growth; linear models miss interactions.")
    R.h2("3.9 Chapter Conclusion")
    P(f"PBL 1 delivered a clean store-week dataset and a baseline that cuts the error by {1 - lin.MAE / nav.MAE:.0%} "
      f"(MAE {lin.MAE:.2f}, R² {lin['R²']:.2f}), with Bayesian intervals for safety stock; Chapter 5 reuses its setup.")

    # ================================================================== CHAPTER 4
    R.chapter("Unsupervised Learning and Sequential Data Modelling – (Units 3 and 4)")
    R.h2("4.1 Overview of Allotted Unit 3 and Unit 4 Topics")
    P("**Unit 3 — clustering and dimensionality reduction:** K-Means [@lloyd] groups the 58 dark stores by demand profile, "
      "and PCA projects the standardised features to two dimensions for visualisation [@esl]. The question is *where*: "
      "which stores can share a stocking and staffing plan? **Unit 4 — sequential modelling:** a Gaussian hidden Markov "
      "model [@rabiner], trained with Baum–Welch and decoded with Viterbi, finds the hidden demand regimes of the whole "
      "network over 638 days. The question is *when*: which regime is the network in, how long does it last and what "
      "comes next? Both units use the original dataset; no separate dataset is used.")

    R.h2("4.2 Unit 3 - Dataset Description and Preprocessing")
    P("The dataset and its preprocessing are described in Sections 3.2–3.3. For Unit 3 the clean orders were "
      "aggregated to one row per store with 20 candidate features: orders per day, night (10 pm–6 am) and morning "
      "(6–10 am) shares, the share of baskets containing each food category, basket value, items, discount share and "
      "service indicators. Each store rate is estimated from about 2,000 orders (1,955 on average) and varies slightly by chance, so "
      "features were selected by their **signal-to-noise ratio** — the between-store variance divided by the variance "
      "expected from sampling noise (Eq. 4.1; SNR ≈ 1 means pure chance). The selected features were standardised "
      "with Eq. (3.3) so that none dominates the Euclidean distance.")
    R.eq(r"\text{SNR}_f = \operatorname{Var}_s(\bar{x}_{s,f}) \Big/ \frac{1}{S}\sum_{s=1}^{S} \frac{\sigma^2_{s,f}}{n_s}")

    R.h2("4.3 Unit 3 – Theory, EDA, Implementation and Results")
    P("The EDA of Chapter 3 is not repeated. The additional feature analysis is shown in Figure 4.1.")
    P(f"**Interpretation of {R.next_fig()}.** " "Demand volume and time-of-day shares differ between stores far more than chance (SNR 236, 32 "
      "and 17) and the food-category shares four to seven times. Ten **demand** features were selected: orders per day, "
      "night and morning shares, the shares of six food categories and basket value (weak overall, but it separates "
      "the office stores). Cancellation, complaint, items and distance are pure noise; discount share follows store age "
      "(r = −0.92 with opening date); service outcomes describe daily conditions, not stores, and are used only to "
      "describe the clusters.")
    R.fig(F("02_clustering", "feature_selection.png"), "Signal-to-noise ratio of the 20 candidate store features (blue = used for clustering)", 5.2)
    P(f"**Implementation.** K-Means was run for *k* = 2–8 with 50 k-means++ initialisations and seed 42 "
      f"({R.next_tab()}, Figure 4.2).")
    ks = D["ksel"]
    R.table("Choice of the number of clusters k",
            ["k", "WCSS", "Silhouette ↑", "DB ↓", "CH ↑", "Cluster sizes"],
            [[str(int(r.k)), f"{r['inertia (WCSS)']:.1f}", f"{r['silhouette ↑']:.3f}", f"{r['Davies–Bouldin ↓']:.3f}",
              f"{r['Calinski–Harabasz ↑']:.1f}", r["cluster sizes"].strip("[]")] for _, r in ks.iterrows()],
            [0.4, 0.8, 1.05, 0.8, 0.8, 1.91])
    P(f"**Choice of k = 3 ({R.next_fig()}).**" " k = 2 scores best but only separates the eleven student-hub stores. k = 3 is second best on "
      "every criterion, the elbow bends there, and it separates a third, business-relevant group (office stores); beyond "
      f"k = 3 all criteria worsen and clusters of one to four stores appear. The final model fitted in "
      f"{km['fit_time_ms']:.0f} ms and gave clusters of 42, 11 and 5 stores, named from their centres (Figure {R.chap}.{R.n_fig + 2}).")
    R.fig(F("02_clustering", "k_selection.png"), "Elbow method and silhouette score for k = 2 to 8 (dashed line = chosen k)")
    P(f"**Interpretation of {R.next_fig()}.** " "PC1 and PC2 explain 78% of the variation, so the 2-D view is faithful. The late-night snack hubs "
      "form a distant, compact group (night share +1.9 standard deviations, morning share −1.9); the daytime office "
      "stores a small group with high fresh food, staples and basket value; the family residential stores the large "
      "central group with the strongest morning share.")
    R.fig(F("02_clustering", "clusters_pca_and_profiles.png"), "Dark stores in the space of the first two principal components and standardised cluster profiles")

    R.h2("4.4 Unit 3 – Algorithm Theory and Mathematical Formulation")
    P("Given standardised store vectors **x**_{1}, …, **x**_{n} (*n* = 58, *d* = 10), K-Means [@lloyd] finds *k* clusters "
      "*C*_{j} with centres **μ**_{j} that minimise the within-cluster sum of squares, alternating an assignment step and "
      "an update step until assignments stop changing:")
    R.eq(r"J = \sum_{j=1}^{k} \sum_{\mathbf{x}_i \in C_j} \lVert \mathbf{x}_i - \boldsymbol{\mu}_j \rVert^2, \qquad c_i = \arg\min_{j} \lVert \mathbf{x}_i - \boldsymbol{\mu}_j \rVert^2, \qquad \boldsymbol{\mu}_j = \frac{1}{\left| C_j \right|}\sum_{\mathbf{x}_i \in C_j} \mathbf{x}_i")
    P("Each step can only decrease *J*, so the algorithm converges to a (local) minimum — hence 50 restarts; it assumes "
      "compact, roughly spherical clusters. PCA projects the data on the leading eigenvectors of its covariance matrix "
      "[@esl]. With *a*(*i*) the mean distance of store *i* to its own cluster and *b*(*i*) to the nearest other "
      "cluster, the silhouette [@rousseeuw], Davies–Bouldin and Calinski–Harabasz indices are")
    R.eq(r"s(i) = \frac{b(i) - a(i)}{\max\{a(i),\, b(i)\}}, \qquad \text{DB} = \frac{1}{k}\sum_{j=1}^{k} \max_{l \neq j} \frac{\sigma_j + \sigma_l}{\lVert \boldsymbol{\mu}_j - \boldsymbol{\mu}_l \rVert}, \qquad \text{CH} = \frac{\operatorname{tr}(\mathbf{B}_k)/(k-1)}{\operatorname{tr}(\mathbf{W}_k)/(n-k)}")
    P("where σ_{j} is the mean distance of cluster *j*'s members to its centre and **B**_{k}, **W**_{k} the between- and "
      "within-cluster scatter matrices; silhouette and CH should be high, DB low. The adjusted Rand index (ARI, 1 = "
      "identical partitions) measures stability and agreement with Ward clustering and a Gaussian mixture. K-Means fits "
      "because each centre is an average store profile that translates directly into a stocking plan.")

    R.h2("4.5 Unit 3 Proposed System Architecture / Workflow")
    P(f"{R.next_fig()} shows the workflow. " "Store features are built (1–2), filtered by SNR (3) and standardised (4); k is chosen with four criteria (5), the "
      "final model is fitted (6), validated (7) and interpreted (8), giving a cluster label per store (9).")
    R.fig(os.path.join(HERE, "fig_workflow_unit3.png"), "Workflow of Unit 3 — clustering of dark stores")

    R.h2("4.6 Unit 3 – Results and Performance Evaluation")
    P("Table 4.2 gives the mean profile of each cluster and Table 4.3 the location type of its stores.", keep_next=True)
    pr = D["prof"]
    lab = {"orders_per_day": "Orders per day", "night_share": "Night orders (%)", "morning_share": "Morning orders (%)",
           "share_fruit_veg": "Baskets with fruit & veg (%)", "share_snacks_bev": "Baskets with snacks & bev. (%)",
           "share_dairy_bread_eggs": "Baskets with dairy & bread (%)", "share_atta_rice_dal": "Baskets with staples (%)",
           "share_icecream_desserts": "Baskets with desserts (%)", "share_instant_frozen": "Baskets with instant food (%)",
           "avg_basket_value": "Average basket value (₹)", "late_rate": "Late deliveries (%) *", "complaint_rate": "Complaints (%) *",
           "stockout_rate": "Stock-out orders (%) *"}
    R.table("Cluster profiles (mean of the stores in each cluster; * = not used for clustering)",
            ["Feature", "Family residential (42)", "Late-night snack hubs (11)", "Daytime office (5)"],
            [[lab[i]] + [f"{v:.1f}" for v in r.values] for i, r in pr.iterrows()],
            [2.35, 1.15, 1.15, 1.11])
    cl = D["cloc"]
    R.table("Location type of the stores in each cluster (location was not used for clustering)",
            ["Cluster", "Office", "Premium Res.", "Residential", "Student / PG"],
            [[i] + [str(v) for v in r.values] for i, r in cl.iterrows()], [2.2, 0.8, 1.0, 0.9, 0.86])
    P(f"**Interpretation of {R.next_fig()}.** " "The clusters match location exactly although location was not used:")
    B(["**Family residential stores** (42, residential areas): strongest morning peak (18.4% of orders at 6–10 am); dairy "
       "in 58% and fresh produce in 47% of baskets → stock milk, bread and produce before 6 am; morning and evening shifts.",
       "**Late-night snack hubs** (11, student / PG hubs): 37.7% of orders at night; most snacks (69% of baskets), "
       "desserts and frozen food → night stock of snacks, cold drinks and frozen food; late rider shifts.",
       "**Daytime office stores** (5, tech parks): fewest night orders (15.2%), a 9 am peak, the highest basket value "
       "(₹459) and most fresh food and staples → weekday daytime stock, light night staffing."])
    R.fig(F("02_clustering", "cluster_baskets_and_hours.png"), "Food categories in the basket and hourly order pattern of each cluster")
    P("Late, complaint and stock-out rates hardly differ between clusters (27–28%, about 5% and 12–13%): service quality "
      "depends on daily conditions rather than store type, which motivates Unit 4. The evaluation of the clustering is "
      "summarised in Table 4.5 (Section 4.11).")

    R.h2("4.7 Unit 4 – Sequential Dataset Description")
    P(f"The Unit 4 sequence was derived from the original dataset by aggregating the 1,13,390 clean food orders of all "
      f"stores per day ({R.next_tab()}).", keep_next=True)
    R.table("Description of the Unit 4 sequential dataset",
            ["Property", "Value"],
            [["Source", "Derived from the original Kaggle dataset; dataset/daily_series.csv"],
             ["Timestamps / sampling", "Calendar days 01-01-2025 to 30-09-2026, one observation per day"],
             ["Sequence / length", "One sequence (whole network of 58 stores); 638 days (365 + 273)"],
             ["Missing timestamps", "None — 637 consecutive one-day steps"],
             ["Observations", "Demand index (Eq. 4.4) and late-delivery rate"],
             ["Hidden state", "Demand regime — unlabelled; 3 states learned"],
             ["Not used in training", "Rain share and festival name — used only for interpretation and validation"],
             ["Limitations", "Single network-level sequence; two monsoon seasons; noisy early weeks of 2025"]],
            [1.65, 4.11], align=[L, L], bold_first=True)

    R.h2("4.8 Unit 4 – Sequential Preprocessing and EDA")
    P("Temporal order was preserved throughout. Daily orders grow about five-fold (59.7 per day in 2025 Q1, 294.5 in "
      "2026 Q3) and follow a weekly rhythm, so a model fitted to raw counts would only learn \"early\" and \"late\" "
      "periods. Growth and weekday effects were removed:")
    R.eq(r"T_t = \operatorname{median}\{y_{t-14}, \ldots, y_{t+14}\}, \qquad \hat{y}_t = T_t\, f_{\text{dow}(t)}, \qquad \text{DI}_t = \log\frac{y_t}{\hat{y}_t}")
    P("where *T*_{t} is a centred 29-day rolling median, *f*_{d} the weekday factor estimated on regular days (weekend "
      "1.12, weekdays 0.93–0.96) and DI_{t} the **demand index** (0 = normal day, +0.3 ≈ 35% above normal). The second "
      "observation is the **late-delivery rate**. Both are on comparable small scales (DI −0.52 to 0.77, late rate 0.11 "
      "to 0.47) and were not rescaled. For out-of-sample testing the sequence was split chronologically: 2025 for "
      f"training and 2026 for testing.")
    P(f"**Interpretation of {R.next_fig()}.** " "The expected level follows growth and the weekend rhythm, so the demand index is centred on "
      "zero (mean 0.016, std 0.131) with sharp festival spikes (Diwali +0.77). The late rate (Figure 4.9) rises from about "
      "22% to over 30% every June–October. Short demand spikes and long periods of high lateness are exactly the kind of "
      "hidden regimes an HMM can separate.")
    R.fig(F("01_dataset", "daily_demand_index.png"), "Daily food orders with the expected level (top) and the demand index (bottom)", 5.4)

    R.h2("4.9 Unit 4 – Mathematical Formulation")
    P("An HMM [@rabiner] describes observations **o**_{1}, …, **o**_{T} (**o**_{t} = (DI_{t}, late rate_{t}), *T* = 638) "
      "generated by hidden states *q*_{t} ∈ {1, …, *N*}. It is defined by λ = (π, **A**, **B**): initial probabilities "
      "π_{i} = *P*(*q*_{1} = *i*), transition probabilities *a*_{ij} = *P*(*q*_{t+1} = *j* | *q*_{t} = *i*) and Gaussian "
      "emission densities *b*_{i}(**o**_{t}) = N(**o**_{t} | **μ**_{i}, **Σ**_{i}) with diagonal covariance. Under the "
      "first-order Markov and conditional-independence assumptions:")
    R.eq(r"P(\mathbf{O}, \mathbf{Q} \mid \lambda) = \pi_{q_1} b_{q_1}(\mathbf{o}_1) \prod_{t=2}^{T} a_{q_{t-1} q_t}\, b_{q_t}(\mathbf{o}_t), \qquad \alpha_{t+1}(j) = \Big[\sum_{i=1}^{N} \alpha_t(i)\, a_{ij}\Big] b_j(\mathbf{o}_{t+1})")
    P("The forward variable α_{t}(*i*) (and the backward variable β_{t}(*i*)) give the likelihood *P*(**O** | λ) = Σ_{i} "
      "α_{T}(*i*). **Baum–Welch (EM)** computes the posteriors and re-estimates the parameters:")
    R.eq(r"\gamma_t(i) = \frac{\alpha_t(i)\,\beta_t(i)}{P(\mathbf{O} \mid \lambda)}, \quad \xi_t(i,j) = \frac{\alpha_t(i)\, a_{ij}\, b_j(\mathbf{o}_{t+1})\, \beta_{t+1}(j)}{P(\mathbf{O} \mid \lambda)}, \quad \hat{a}_{ij} = \frac{\sum_{t} \xi_t(i,j)}{\sum_{t} \gamma_t(i)}, \quad \hat{\boldsymbol{\mu}}_i = \frac{\sum_{t} \gamma_t(i)\, \mathbf{o}_t}{\sum_{t} \gamma_t(i)}")
    P("EM never decreases the likelihood but finds a local maximum, so 20 random restarts were run. **Viterbi** decoding "
      "finds the most likely state sequence by dynamic programming, δ_{t+1}(*j*) = max_{i}[δ_{t}(*i*) *a*_{ij}] "
      "*b*_{j}(**o**_{t+1}). The number of states was chosen by BIC [@schwarz] with *k* free parameters; the expected "
      "stay in a state and tomorrow's regime probabilities are")
    R.eq(r"\text{BIC} = k \ln T - 2 \ln \hat{L}, \qquad E[\text{duration}_i] = \frac{1}{1 - a_{ii}}, \qquad P(q_{T+1} = j \mid \mathbf{O}) = \sum_{i} \gamma_T(i)\, a_{ij}")
    P("The inference goals are decoding (the regime of every day), persistence (transitions) and prediction (next day); "
      "the model assumes Gaussian observations within a regime and parameters that do not change over time.")

    R.h2("4.10 Unit 4 Proposed System Architecture / Workflow")
    P(f"{R.next_fig()} shows the workflow. " "Daily observations are built (1–3); HMMs with two to five states are trained with Baum–Welch (4) and chosen by "
      "BIC (5); the learned π, **A**, **B** (6) drive Viterbi decoding (7), which is validated on festivals, unseen 2026 "
      "days and time-blind methods (8) and yields daily regimes and a next-day forecast (9).")
    R.fig(os.path.join(HERE, "fig_workflow_unit4.png"), "Workflow of Unit 4 — hidden Markov model of daily demand regimes")

    R.h2("4.11 Results and Performance Evaluation")
    R.h3("4.11.1 Unit 3 evaluation")
    cm = D["cmet"]
    R.table("Unit 3 evaluation — K-Means clustering of the 58 dark stores (k = 3)",
            ["Criterion", "K-Means", "Ward", "GMM"],
            [["Silhouette ↑", f"{cm.iloc[0, 0]:.3f}", f"{cm.iloc[0, 1]:.3f}", f"{cm.iloc[0, 2]:.3f}"],
             ["Davies–Bouldin ↓", f"{cm.iloc[1, 0]:.3f}", f"{cm.iloc[1, 1]:.3f}", f"{cm.iloc[1, 2]:.3f}"],
             ["Calinski–Harabasz ↑", f"{cm.iloc[2, 0]:.2f}", f"{cm.iloc[2, 1]:.2f}", f"{cm.iloc[2, 2]:.2f}"],
             ["Agreement with K-Means (ARI)", "1.000", f"{cm.iloc[3, 1]:.3f}", f"{cm.iloc[3, 2]:.3f}"],
             ["Stability over 20 restarts (mean ARI)", f"{km['stability_ari']:.2f}", "—", "—"],
             ["Stores with negative silhouette", "0 of 58", "—", "—"],
             ["Match with location type", "100%", "—", "—"]],
            [2.7, 1.0, 1.0, 1.06])
    P("A silhouette of 0.42 indicates a reasonable structure (above 0.25 reasonable, above 0.5 strong) and no store is "
      "closer to another cluster. Ward clustering agrees almost completely (ARI 0.94), the Gaussian mixture gives the "
      "same partition, and restarts reproduce the solution (ARI 0.85). Without basket value the silhouette drops to "
      "0.268 and the office stores scatter, confirming that this feature is needed.")
    R.h3("4.11.2 Unit 4 results")
    ns = D["nsel"]
    R.table("Selection of the number of hidden states",
            ["States", "Log-likelihood", "Parameters", "AIC ↓", "BIC ↓"],
            [[str(int(r.states)), f"{r['log-likelihood']:.1f}", str(int(r.parameters)), f"{r['AIC ↓']:.1f}", f"{r['BIC ↓']:.1f}"]
             for _, r in ns.iterrows()], [0.9, 1.3, 1.1, 1.2, 1.26])
    P("In Table 4.6, BIC is lowest for **three states** (−2,802.2); AIC keeps falling because its penalty is weaker. A model on the "
      "demand index alone split ordinary days by calendar quarter (a size artefact of the small early business), which "
      "is why the late rate was added. The learned initial probabilities are π = (1, 0, 0): with a single sequence, π "
      f"only records that the first day was normal. {R.next_tab()} and {R.next_fig()} show the transitions and emissions.")
    tr, em = D["trans"], D["emis"]
    stay = [1 / (1 - tr.values[i, i]) for i in range(3)]
    R.table("Transition matrix A, expected stay and Gaussian emission means (std) of each state",
            ["From \\ to", "Normal", "Monsoon", "Festival", "Stay (days)", "Demand index", "Late rate"],
            [[s, *[f"{v:.3f}" for v in tr.values[k]], f"{stay[k]:.1f}",
              f"{em.iloc[k]['demand index mean']:+.3f} ({em.iloc[k]['demand index std']:.3f})",
              f"{em.iloc[k]['late rate mean']:.3f} ({em.iloc[k]['late rate std']:.3f})"]
             for k, s in enumerate(["Normal", "Monsoon / disrupted", "Festival peak"])],
            [1.25, 0.7, 0.8, 0.75, 0.6, 0.9, 0.76])
    P(f"**Interpretation of {R.next_fig()}.** " "*Normal* and *Monsoon / disrupted* are very persistent (stay probabilities 0.972 and 0.981; "
      "about 35 and 54 days) and have the same demand distribution; they differ in lateness (22.5% vs 32.3% late, and a "
      "0.67 probability of more than 30% late on a monsoon day). *Festival peak* has 32% higher demand with a wide spread, "
      "is short-lived (about 2.4 days) and usually returns to Normal (0.279).")
    R.fig(F("03_hmm", "transitions_and_emissions.png"), "Transition probabilities and the emission densities of the demand index", 5.0)
    R.fig(F("03_hmm", "decoded_regimes.png"), "Decoded regime of every day (Viterbi): demand index (top) and late deliveries (bottom)", 5.3)
    dl, rgs = D["daily"], D["regime"]
    days = dl.state.value_counts()
    runs = dl.groupby((dl.state != dl.state.shift()).cumsum()).state.agg(["first", "size"]).groupby("first")["size"].mean()
    R.table("Summary of the decoded regimes",
            ["Regime", "Days", "Share", "Mean run", "Orders / day", "Late rate", "Rain share", "Months"],
            [[s, str(days[s]), f"{days[s] / len(dl):.1%}", f"{runs[s]:.1f}", f"{rgs.loc[s, 'orders']:.0f}",
              f"{rgs.loc[s, 'late_rate']:.1%}", f"{rgs.loc[s, 'rain_share']:.1%}", mo]
             for s, mo in zip(["Normal", "Monsoon / disrupted", "Festival peak"], ["Nov–May", "Jun–Oct", "festivals"])],
            [1.45, 0.5, 0.6, 0.6, 0.65, 0.6, 0.6, 0.76])
    fd = D["fest"]
    grp = fd.groupby("event", sort=False).agg(days=("state", "size"), peak=("state", lambda s: (s == "Festival peak").sum()),
                                               di=("demand_index", "mean"))
    R.table("Festival days and how many were decoded as Festival peak",
            ["Event", "Festival days", "Decoded as peak", "Mean demand index"],
            [[e, str(r.days), str(r.peak), f"{r.di:+.3f}"] for e, r in grp.iterrows()] +
            [["**Total**", f"**{len(fd)}**", f"**{int(grp.peak.sum())}**", ""]],
            [2.0, 1.2, 1.3, 1.26])
    P(f"**Interpretation of Tables 4.8–4.9.** Without being told about rain or festivals, the model finds a Normal regime (November–May), "
      f"a Monsoon / disrupted regime (June–October; half the orders placed in rain, a third of deliveries late) and short "
      f"Festival peaks. There are only {h['switches']} switches in 638 days and the mean posterior certainty is "
      f"{h['mean_certainty']:.1%}. 18 of 22 festival days (**{h['festival_recall']:.0%}**) are decoded as peaks; the misses "
      f"are festivals with a mild rise (Christmas, Valentine's Day 2026, New Year's Day 2025), and only {h['false_peaks']} "
      "regular days (09–13 January 2025, the noisy first weeks) are false peaks.")
    oos, sm = D["oos"], D["same"]
    nx = D["nxt"]
    P(f"**Out-of-sample test.** Trained on 2025 and scored on the unseen 2026 days, the HMM has a log-likelihood of "
      f"{oos.iloc[0, 0]:.2f} per day, against {oos.iloc[1, 0]:.2f} for the same states without transitions (Gaussian "
      f"mixture) and {oos.iloc[2, 0]:.2f} for a single Gaussian: knowing yesterday's regime makes today more "
      f"predictable. On 30-09-2026 the network is in the monsoon regime (probability {nx.iloc[1, 0]:.3f}); tomorrow it "
      f"stays there with probability {nx.iloc[1, 1]:.3f}, with a {nx.iloc[2, 1]:.1%} chance of a festival peak. "
      f"{R.next_tab()} and {R.next_fig()} compare three methods that label the same 638 days.")
    def fv(i, v):
        if pd.isna(v): return "—"
        if i in ("regime switches", "regular days in peak group"): return f"{v:.0f}"
        if "days in" in i: return f"{v:.1%}"
        return f"{v:.3f}" if "silhouette" in i or "log-lik" in i else f"{v:.1f}"
    R.table("Same-data comparison of three methods that label the 638 days",
            ["Criterion", "K-Means", "GMM", "HMM"],
            [[i[0].upper() + i[1:]] + [fv(i, v) for v in r.values] for i, r in sm.iterrows()], [2.9, 0.95, 0.95, 0.96])
    P(f"**Interpretation of {R.next_fig()}.** " "Time-blind methods switch 251–270 times because one rainy-looking day flips them; the HMM "
      "switches 19 times (runs of about 32 days). K-Means catches slightly more festival days only because its peak group "
      "also contains 61 ordinary days. The silhouette favours flickering labels because it ignores time. "
      "Table 4.11 summarises the Unit 4 evaluation.")
    R.fig(F("04_comparison", "label_of_each_day.png"), "Label of each day: methods without time order flicker, the HMM gives stable regimes")
    R.table("Unit 4 evaluation — Gaussian HMM of the daily sequence",
            ["Criterion", "Value", "Assessment"],
            [["Number of states", "3 (lowest BIC)", f"log-likelihood {h['log_likelihood']:,.1f}"],
             ["Regime switches / certainty", f"{h['switches']} / {h['mean_certainty']:.1%}", "stable, confident regimes"],
             ["Festival recall", f"{h['festival_recall']:.0%} (18 of 22)", f"{h['false_peaks']} false peak days"],
             ["Jun–Sep days in disrupted regime", "94.7%", "monsoon found without weather"],
             ["Out-of-sample log-likelihood", f"{h['oos_hmm']:.2f}", f"vs {h['oos_gmm']:.2f} without transitions"],
             ["Training time", f"{h['fit_time_s']:.1f} s (20 restarts)", "negligible"]],
            [2.35, 1.55, 1.86], align=[L, C, L])

    R.h2("4.12 Discussion, Findings and Limitations")
    P("**Unit 3:** stable, location-consistent store types, made possible by SNR feature selection; limited by 58 stores "
      "and a preset *k*. **Unit 4:** monsoon and festival regimes found without weather or festival input and validated "
      "on an unseen year; Normal and Monsoon differ only in lateness, small festivals stay within noise, and Markov "
      "transitions imply geometric durations.")
    R.h2("4.13 Chapter Conclusion")
    P("Unit 3 found three store types with distinct plans (silhouette 0.42) and Unit 4 three demand regimes with 82% "
      "festival recall. Their data units differ, so their outputs are combined (Chapter 6) rather than compared.")

    # ================================================================== CHAPTER 5
    R.chapter("Advanced Predictive Modelling and Ensemble Learning (Unit 5)")
    R.h2("5.1 Overview of Allotted Unit 5 Topics")
    P("**Gradient boosting** [@friedman] was selected as the ensemble method best suited to the tabular store-week data "
      "of PBL 1: it can model the non-linear effects and interactions that linear regression cannot. It builds an "
      "additive model of *M* small **CART regression trees** [@esl], each fitted to the errors of the ensemble so far and "
      "added with a learning rate ν; training each tree on a random 80% of the rows (stochastic boosting) reduces "
      "variance:")
    R.eq(r"F_M(\mathbf{x}) = F_0 + \nu \sum_{m=1}^{M} h_m(\mathbf{x})")
    R.h2("5.2 Dataset Description and Preprocessing")
    P("Unit 5 uses the store-week table of Chapter 3 with the same split (2,531 training and 1,508 test store-weeks) and "
      "metrics, so the results are directly comparable; trees need no scaling. The additional features follow from "
      f"Section 5.3 ({R.next_tab()}).", keep_next=True)
    R.table("Additional features for gradient boosting (growth-free formulation)",
            ["Feature / target", "Definition", "Purpose"],
            [["target *z*", "log(orders / roll4)", "relative change, not level"],
             ["`r1`, `r2`, `r4`", "lag1, lag2, lag4 divided by roll4", "momentum free of store size"],
             ["`woy`", "ISO week of the year", "seasonal position"],
             ["kept", "roll4, store age, festival days, monsoon, area", "level, calendar, location"]],
            [1.35, 2.5, 1.91], align=[L, L, L])
    R.h2("5.3 EDA and Feature Analysis")
    P("A regression tree can only predict values within the range of its training targets, so it cannot "
      f"**extrapolate** growth. {R.next_fig()} examines this.")
    P(f"**Interpretation of {R.next_fig()}.** " "Many stores lie near or above the diagonal: **19.2% of test store-weeks exceed the store's "
      "training maximum**, which a tree trained on order levels cannot predict. The growth-free target log(orders / roll4) "
      "is centred near zero (mean +0.011, std 0.270, skew −0.40) and has the same range in training and test, because "
      "growth has been divided out.")
    R.fig(F("06_boosting", "boosting_feature_analysis.png"), "Per-store training maximum vs test mean, and the distribution of the growth-free target", 5.3)
    R.h2("5.4 Proposed Workflow / Architecture")
    P(f"{R.next_fig()} shows the workflow. " "The PBL 1 data and split (1–2) receive the growth-free target (3); boosted CART trees (4) are tuned with "
      "time-series CV (5) and refitted (6), evaluated (7) and compared with PBL 1 on the same weeks (8–9).")
    R.fig(os.path.join(HERE, "fig_workflow_unit5.png"), "Workflow of Unit 5 — gradient boosting for weekly demand forecasting")
    R.h2("5.5 Algorithm Theory and Mathematical Formulation")
    P("Gradient boosting minimises a loss by functional gradient descent [@friedman]. With squared-error loss it starts "
      "from *F*_{0} = ȳ; at each stage a tree *h*_{m} with *J* leaves (regions *R*_{jm}, values γ_{jm}) is fitted to the "
      "pseudo-residuals and added with shrinkage ν:")
    R.eq(r"r_{im} = -\left[\frac{\partial L(y_i, F(\mathbf{x}_i))}{\partial F(\mathbf{x}_i)}\right]_{F = F_{m-1}} = y_i - F_{m-1}(\mathbf{x}_i), \qquad F_m(\mathbf{x}) = F_{m-1}(\mathbf{x}) + \nu \sum_{j=1}^{J} \gamma_{jm}\, \mathbb{1}(\mathbf{x} \in R_{jm})")
    P("Each CART split is the feature and threshold that most reduce the squared error. Because trees cannot extrapolate, "
      "the model was trained on the relative change and transformed back:")
    R.eq(r"z_t = \log\frac{y_t}{\text{roll4}_t}, \qquad \hat{y}_t = \text{roll4}_t \cdot \exp\big(\hat{F}(\mathbf{x}_t)\big)")
    P("**Settings and model selection.** GradientBoostingRegressor [@sklearn] with squared-error loss, subsample 0.8 and "
      "seed 42; grid *n*_estimators ∈ {200, 400, 800}, learning rate ∈ {0.02, 0.05}, max_depth ∈ {2, 3} (12 combinations), "
      "scored by MAE with 5-fold `TimeSeriesSplit`; the best setting was refitted on all training weeks. Feature "
      "importance was measured by impurity reduction and checked by permutation on the test weeks.")
    R.h2("5.6 Results and Performance Evaluation")
    R.table("Gradient boosting on the original features vs the growth-free target (test weeks)",
            ["Version", "Best setting (CV)", "MAE", "RMSE", "R²", "WAPE"],
            [["1 — original features", "200 trees, depth 2, ν 0.05", f"{gbr.MAE:.3f}", f"{gbr.RMSE:.3f}", f"{gbr['R²']:.3f}", f"{gbr.WAPE:.1%}"],
             ["2 — growth-free target", "200 trees, depth 2, ν 0.02", f"{gbg.MAE:.3f}", f"{gbg.RMSE:.3f}", f"{gbg['R²']:.3f}", f"{gbg.WAPE:.1%}"]],
            [1.6, 1.85, 0.55, 0.6, 0.55, 0.61], align=[L, L, C, C, C, C])
    P("Table 5.2 shows that plain boosting is worse than linear regression (MAE 5.92 vs 5.35) because it cannot predict the 19% of test weeks "
      "above the training range; the growth-free version cuts MAE by 10% to 5.32. Cross-validation favoured small "
      "models (200–400 trees, depth 2).")
    P(f"**Interpretation of {R.next_fig()}.** " "Festival days (0.34) and `roll4` (0.21) are the most important features, followed by `r4`, the "
      "week of the year and store age; permutation importance on the test weeks confirms only festival days and `roll4`, "
      "so the other splits mostly fit training noise. The validation error (last 13 training weeks) reaches its minimum "
      "of 4.70 at 133 trees and rises to 4.87 at 800 while the training error keeps falling — over-fitting that the "
      "cross-validated choice of 200 trees avoids.")
    R.fig(F("06_boosting", "importance_and_learning_curve.png"), "Impurity-based feature importance of the final model and the learning curve", 5.3)
    R.h2("5.7 Comparison with Previous Models")
    R.table("Comparison of all forecasting models on the same 1,508 test store-weeks",
            ["Model", "Unit", "MAE ↓", "RMSE ↓", "R² ↑", "WAPE ↓"],
            [[i, u, f"{r.MAE:.3f}", f"{r.RMSE:.3f}", f"{r['R²']:.3f}", f"{r.WAPE:.1%}"]
             for (i, r), u in zip(comp.iterrows(), ["base", "1", "2", "2", "5", "5"])],
            [2.5, 0.5, 0.68, 0.75, 0.6, 0.73])
    P(f"**Improvements and declines.** Growth-free gradient boosting is the most accurate model on every metric (MAE "
      f"{gbg.MAE:.3f}, R² {gbg['R²']:.3f}), but its gain over linear regression is only {1 - gbg.MAE / lin.MAE:.1%}; "
      "boosting on the original features is second worst. If weekly orders were Poisson around a perfect forecast, the "
      f"MAE would still be about **{D['gb']['noise_floor_mae']:.2f}** (noise floor): the best models are within about one "
      "order of this limit, so most remaining error is unpredictable week-to-week variation.")
    R.h2("5.8 Discussion, Findings and Limitations")
    P(f"Boosting generalises once the target is stationary, but gains little because the lag signal is linear and near "
      f"the noise floor; it is harder to interpret, slower to tune ({D['gb']['search_time_s']:.0f} s) and has no interval. "
      f"**Recommendation:** Bayesian regression for weekly replenishment; boosting once richer features are added.")
    R.h2("5.9 Chapter Conclusion")
    P(f"Unit 5 showed that trees must be given a growth-free target; the tuned ensemble is the most accurate model "
      f"(MAE {gbg.MAE:.2f}, R² {gbg['R²']:.3f}), and its small margin shows PBL 1 was near the feature limit.")

    # ================================================================== CHAPTER 6
    R.chapter("Integrated Discussion and Findings")
    R.table("Summary of the five units",
            ["Unit", "Method", "Key result", "Supply-chain use"],
            [["1", "Linear Regression", f"MAE {lin.MAE:.2f} (naive {nav.MAE:.2f})", "replenishment"],
             ["2", "Ridge, Bayesian", "same MAE; 90.6% coverage", "safety stock"],
             ["3", "K-Means (+ PCA)", "3 store types; silhouette 0.42", "store plans"],
             ["4", "Gaussian HMM", "3 regimes; festival recall 82%", "regime buffers"],
             ["5", "Gradient Boosting", f"best MAE {gbg.MAE:.2f}", "best forecast"]],
            [0.5, 1.6, 2.26, 1.4], align=[C, L, L, L])
    P("Table 6.1 summarises how the units answer *how much* (PBL 1, 3), *where* (Unit 3) and *when* (Unit 4). Units 1, 2 "
      "and 5 share task, test weeks and metrics and are compared directly (Table 5.3); Units 3 and 4 use different data "
      f"units, so their outputs were combined by labelling every order with its store's cluster and day's regime "
      f"({R.next_fig()}). The critical findings follow.")
    R.fig(F("04_comparison", "late_by_cluster_and_regime.png"), "Late-delivery rate by store cluster (K-Means) and demand regime (HMM)", 3.6)
    R.numbered([cite(x) for x in [
        "**Recent demand drives next week's orders;** all forecasts are near the noise floor, so new information "
        "(prices, promotions, weather) matters more than model complexity.",
        "**Simple models are competitive:** linear and Bayesian regression are within 0.7% of the best ensemble.",
        "**Stores differ in what and when, not in service quality.**",
        "**The regime matters more than the store type for delays:** monsoon adds 9–11 points of late deliveries in "
        "every cluster.",
        "**Time order adds value:** the HMM gives stable regimes and predicts unseen days best. Main limitations: "
        "simulated data, 58 stores, two monsoon seasons and no price data."]])

    # ================================================================== CHAPTER 7
    R.chapter("Overall Conclusion and Future Scope")
    R.h2("7.1 Conclusion")
    P("This project built an integrated machine learning pipeline for food supply planning in Indian quick commerce from "
      "one public dataset of 1,95,148 orders. Its achievements are:")
    B(["**Data (Unit 1):** a reproducible pipeline producing 1,13,390 clean food orders, a 4,039-row store-week table "
       "and a 638-day sequence.",
       f"**Forecasting (Units 1–2):** error cut by {1 - lin.MAE / nav.MAE:.0%} versus the naive forecast (MAE "
       f"{lin.MAE:.2f}, R² {lin['R²']:.2f}), with Bayesian intervals for safety stock.",
       "**Segmentation (Unit 3):** three stable store types (silhouette 0.42, stability 0.85) with distinct plans.",
       "**Regimes (Unit 4):** normal, monsoon and festival regimes found without labels; 82% festival recall.",
       f"**Ensemble (Unit 5):** growth-free boosting is the most accurate model (MAE {gbg.MAE:.2f}), close to the noise "
       "floor."])
    P("The models answer how much to send, to which type of store and under which conditions, and show that careful "
      "data preparation and validation matter more than model complexity.")
    R.h2("7.2 Future Scope")
    B(["**Data:** real orders with product-level quantities, prices, promotions and inventory.",
       "**Features and tuning:** HMM regime probabilities, weather and promotion calendars as forecast features; count "
       "models and quantile boosting for calibrated intervals.",
       "**Sequential models:** HMMs per city or cluster, input–output HMMs and semi-Markov durations.",
       "**Optimisation engine:** use the forecasts, intervals, store types and regimes as inputs to safety-stock "
       "[@achkar] and learned replenishment policies [@geevers, @rizqi] with shelf-life limits [@claassen].",
       "**Deployment:** weekly retraining, regime alerts for store managers, error monitoring and SHAP explanations."])

    # ================================================================== REFERENCES
    R.back_heading("REFERENCES")
    for i, k in enumerate(ORDER, 1):
        p = R.doc.add_paragraph()
        fmt(p, L, ls=1.5, after=4, hanging=0.45, right=0)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(0.45))
        add_rich(p, f"[{i}]\t" + REFS[k], size=12)
    missing = [k for k in REFS if k not in ORDER]
    assert not missing, missing

    # ================================================================== APPENDIX A
    R.back_heading("APPENDIX A", toc_title="APPENDIX A – CODING")
    p = R.doc.add_paragraph(); fmt(p, C, ls=1.0, after=10)
    set_font(p.add_run("CODING"), 16, bold=True, color=RGBColor(0, 0, 0))
    P("The implementation consists of six Jupyter notebooks run in order (01 → 06). The listing below gives the core "
      "implementation of every unit — data preparation, model configuration, training and evaluation — with the same "
      "settings as the notebooks (numpy, pandas, scikit-learn and hmmlearn; SEED = 42); imports, loading, plotting and "
      "file-saving statements are omitted for length. The complete "
      "notebooks with all outputs are available in the project files.")
    for head, link, code in [
            ("A.1 Chapter 3 – Implementation Methodology and Algorithm Configuration",
             "01_Create_Dataset.ipynb, 05_PBL1_Regression.ipynb", APX_CH2),
            ("A.2 Chapter 4 – Unit 3 – Proposed Workflow and Implementation Methodology",
             "02_Clustering_KMeans.ipynb", APX_U3),
            ("A.3 Chapter 4 – Unit 4 – Proposed Workflow and Implementation Methodology",
             "03_HMM.ipynb, 04_Comparison_and_Results.ipynb", APX_U4),
            ("A.4 Chapter 5 – Implementation Methodology and Hyperparameter Configuration",
             "06_PBL3_Gradient_Boosting.ipynb", APX_U5)]:
        R.h2(head)
        R.code(code, after=4)
ORDER = []


def cite(text):
    def rep(m):
        keys = [k.strip().lstrip("@") for k in m.group(1).split(",")]
        nums = []
        for k in keys:
            assert k in REFS, k
            if k not in ORDER:
                ORDER.append(k)
            nums.append(ORDER.index(k) + 1)
        return "[" + "], [".join(map(str, nums)) + "]"
    return re.sub(r"\[(@[\w@, ]+)\]", rep, text)


def n(x, d=2):
    return f"{x:,.{d}f}"


def inr(x):
    s = f"{abs(x):.0f}"; head, tail = s[:-3], s[-3:]; parts = []
    while len(head) > 2:
        parts.insert(0, head[-2:]); head = head[:-2]
    if head: parts.insert(0, head)
    return ("-" if x < 0 else "") + ",".join(parts + [tail])


def load():
    D = {}
    D["metrics"] = pd.read_csv(F("05_regression", "metrics.csv"), index_col=0)
    D["coef"] = pd.read_csv(F("05_regression", "coefficients.csv"), index_col=0)
    D["reg"] = json.load(open(F("05_regression", "summary.json")))
    D["gb"] = json.load(open(F("06_boosting", "summary.json")))
    D["comp"] = pd.read_csv(F("06_boosting", "comparison_with_pbl1.csv"), index_col=0)
    D["cv"] = pd.read_csv(F("06_boosting", "cv_results.csv"))
    D["imp"] = pd.read_csv(F("06_boosting", "feature_importance.csv"), index_col=0)
    D["snr"] = pd.read_csv(F("02_clustering", "feature_signal_to_noise.csv"), index_col=0)
    D["ksel"] = pd.read_csv(F("02_clustering", "k_selection.csv"))
    D["prof"] = pd.read_csv(F("02_clustering", "cluster_profiles.csv"), index_col=0)
    # recompute the means at full precision (the CSV is rounded to 2 decimals; rounding twice can change the 1st decimal)
    _st = pd.read_csv(os.path.join(PROJ, "dataset", "store_features.csv"), index_col=0)
    _cl = pd.read_csv(F("02_clustering", "store_clusters.csv"), index_col=0)
    D["prof"] = (_st.loc[_cl.index, D["prof"].index].groupby(_cl["cluster"]).mean().T)[D["prof"].columns]
    D["cloc"] = pd.read_csv(F("02_clustering", "cluster_vs_location.csv"), index_col=0)
    D["cmet"] = pd.read_csv(F("02_clustering", "cluster_metrics.csv"), index_col=0)
    D["km"] = json.load(open(F("02_clustering", "summary.json")))
    D["hmm"] = json.load(open(F("03_hmm", "summary.json")))
    D["nsel"] = pd.read_csv(F("03_hmm", "state_selection.csv"))
    D["trans"] = pd.read_csv(F("03_hmm", "transition_matrix.csv"), index_col=0)
    D["emis"] = pd.read_csv(F("03_hmm", "emission_parameters.csv"), index_col=0)
    D["eprob"] = pd.read_csv(F("03_hmm", "emission_probabilities.csv"), index_col=0)
    D["fest"] = pd.read_csv(F("03_hmm", "festival_days.csv"))
    D["oos"] = pd.read_csv(F("03_hmm", "out_of_sample.csv"), index_col=0)
    D["nxt"] = pd.read_csv(F("03_hmm", "next_day_forecast.csv"), index_col=0)
    # full-precision values from the saved model (the CSVs are rounded to 4 decimals; rounding twice can shift the
    # third decimal, e.g. 0.02349 -> 0.0235 -> 0.024), so the report shows exactly what notebook 03 prints
    import joblib
    _m = joblib.load(F("03_hmm", "hmm_model.joblib")); _h = _m["hmm"]
    _order = [k for name in ("Normal", "Monsoon / disrupted", "Festival peak") for k, v in _m["state_names"].items() if v == name]
    D["trans"] = pd.DataFrame(_h.transmat_[np.ix_(_order, _order)], index=D["trans"].index, columns=D["trans"].columns)
    _obs = pd.read_csv(os.path.join(PROJ, "dataset", "daily_series.csv"))[list(_m["observations"])].values
    _post = _h.predict_proba(_obs)[-1]
    D["nxt"] = pd.DataFrame({D["nxt"].columns[0]: _post[_order], D["nxt"].columns[1]: (_post @ _h.transmat_)[_order]},
                            index=D["nxt"].index)
    D["regime"] = pd.read_csv(F("03_hmm", "regime_summary.csv"), index_col=0)
    D["daily"] = pd.read_csv(F("03_hmm", "daily_regimes.csv"), parse_dates=["order_date"])
    D["same"] = pd.read_csv(F("04_comparison", "same_data_comparison.csv"), index_col=0)
    D["late"] = pd.read_csv(F("04_comparison", "late_rate_by_cluster_and_regime.csv"), index_col=0)
    D["dict"] = pd.read_csv(os.path.join(PROJ, "dataset", "data_dictionary.csv"))
    D["weekly"] = pd.read_csv(os.path.join(PROJ, "dataset", "store_weekly.csv"), parse_dates=["week"])
    return D



COMMON_SETUP = """# Common setup (first cell of every notebook)
import os, json, time, warnings
import numpy as np, pandas as pd
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression, RidgeCV, BayesianRidge
from sklearn.cluster import KMeans, AgglomerativeClustering
from sklearn.mixture import GaussianMixture
from sklearn.decomposition import PCA
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import TimeSeriesSplit, GridSearchCV
from sklearn.inspection import permutation_importance
from sklearn.metrics import (mean_absolute_error, mean_squared_error, r2_score,
    silhouette_score, silhouette_samples, davies_bouldin_score,
    calinski_harabasz_score, adjusted_rand_score)
from scipy.stats import norm
from hmmlearn.hmm import GaussianHMM
import joblib
warnings.filterwarnings("ignore")
SEED, DATE = 42, "%d-%m-%Y"
"""

_PLOT = re.compile(r"\b(plt|fig|ax|axes|save_fig|save_chart|display|print|im|gs|blues|div|lim|xpos|w|vb|va|"
                   r"ORANGE|BLUE|AQUA|YELLOW|MUTED|INK|INK2|CCOL|SCOL|cols|grid_x|lab|y_pos|order)\b")


def _is_plot(src):
    return bool(_PLOT.search(src))


def _filter(stmts, src):
    """Returns [(node, modified)] keeping modelling statements only."""
    import ast
    out = []
    for s in stmts:
        seg = ast.get_source_segment(src, s)
        if isinstance(s, ast.Expr):
            if not isinstance(s.value, ast.Call) or _is_plot(seg) or re.search(r"\.(head|round|to_frame)\([^)]*\)$", seg):
                continue
            out.append((s, False)); continue
        if isinstance(s, (ast.For, ast.If, ast.With, ast.FunctionDef)) and _is_plot(seg):
            continue                                         # chart-only loops and helpers
        if _is_plot(seg) and not re.search(r"\.to_csv|joblib\.dump|json\.dump", seg):
            continue
        out.append((s, False))
    return out


def core_code(path):
    """Notebook code without the setup cell, charts and print/display statements."""
    import ast
    nb = nbformat.read(path, 4)
    blocks = []
    for k, c in enumerate([c for c in nb.cells if c.cell_type == "code"]):
        src = c.source
        if k == 0:                                           # setup cell: keep only the output folder
            m = re.search(r'^OUT = "[^"]+"', src, re.M)
            blocks.append(m.group(0) + "; os.makedirs(OUT, exist_ok=True)")
            continue
        src = "\n".join("# " + l if l.lstrip().startswith(("%", "!")) else l for l in src.split("\n"))
        tree = ast.parse(src)
        kept = [ast.unparse(n) if mod else ast.get_source_segment(src, n) for n, mod in _filter(tree.body, src)]
        kept = [x for x in kept if not re.match(r"^for f in sorted\(os\.listdir", x)]
        if kept:
            blocks.append("\n".join(kept))
    return "\n\n".join(blocks)


# =====================================================================================================
def build(pages, out):
    ORDER.clear()
    R = Report(TEMPLATE, TITLE)
    R.front_matter(FRONT, os.path.join(HERE, "front_pages_sample.docx"))
    D = load()
    write(R, D)
    R.finish_equations(tempfile.mkdtemp())
    R.sections_and_footers()
    R.build_index_tables(pages)
    from report_engine import normalise
    normalise(R.doc)
    R.doc.core_properties.title = TITLE
    R.doc.core_properties.author = "Project team"
    R.doc.save(out)
    return R


def to_roman(n):
    vals = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    s = ""
    for v, r in vals:
        while n >= v:
            s += r; n -= v
    return s


def find_pages(R, pdf):
    txt = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
    pages = [re.sub(r"\s+", " ", p.replace("ﬁ", "fi").replace("ﬂ", "fl")) for p in txt.split("\f")]
    norm = lambda s: re.sub(r"\s+", " ", s.replace("`", ""))
    main = next(i for i, p in enumerate(pages) if "CHAPTER 1" in p and "1. INTRODUCTION" in p)
    first = next(i for i in range(main) if pages[i].strip().startswith("ABSTRACT"))
    label = lambda i: to_roman(i - first + 5) if i < main else str(i - main + 1)        # ABSTRACT page = v
    out = {}
    for key, word in (("abstract", "ABSTRACT"), ("toc", "TABLE OF CONTENTS"), ("lof", "LIST OF FIGURES"),
                      ("lot", "LIST OF TABLES"), ("abbr", "ABBREVIATIONS")):
        idx = next(i for i in range(first, main) if pages[i].strip().startswith(word))
        out[key] = label(idx)
    for kind, no, title, key in R.toc:
        k = norm(key)
        idx = next((i for i in range(main, len(pages)) if k in pages[i]), None)
        out[key] = label(idx) if idx is not None else "?"
    for kind, items in (("Figure", R.lof), ("Table", R.lot)):
        for num, cap in items:
            k = norm(f"{kind} {num}: {cap}")[:45]
            idx = next((i for i in range(main, len(pages)) if k in pages[i]), None)
            out[f"{kind} {num}:"] = label(idx) if idx is not None else "?"
    return out, len(pages) - 1


if __name__ == "__main__":
    tmp = tempfile.mkdtemp()
    first = os.path.join(tmp, "pass1.docx")
    R = build({}, first)
    soffice = sys.argv[3] if len(sys.argv) > 3 else "soffice"
    subprocess.run(["python", soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, first],
                   check=True, capture_output=True)
    pages, npages = find_pages(R, os.path.join(tmp, "pass1.pdf"))
    print("pages:", npages, "| unresolved:", [k for k, v in pages.items() if v == "?"])
    build(pages, OUT_DOCX)
    print("saved", OUT_DOCX)

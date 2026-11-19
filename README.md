# AI Supply Chain Optimizer — Indian Quick-Commerce Food Demand
21CSC305P Machine Learning Project — PBL 1 (Units 1–2), PBL 2 (Units 3–4) and PBL 3 (Unit 5).

**Dataset:** India Quick Commerce Orders 2025–2026, Kaggle (CC0) —
https://www.kaggle.com/datasets/abirambharathisa/india-quick-commerce-orders-2025-2026
(already included as `data/india_quick_commerce_orders.csv`).

## Folder layout
```
QuickCommerce_ML_Project_Final/
├── Unit_1_2_PBL1_Data_and_Regression/   Units 1–2: data preparation, EDA, Linear / Ridge / Bayesian regression
│   ├── 01_Create_Dataset.ipynb
│   ├── 05_PBL1_Regression.ipynb
│   └── results/01_dataset/, results/05_regression/
├── Unit_3_PBL2_Clustering_KMeans/       Unit 3: K-Means clustering of the dark stores (PCA for visualisation)
│   ├── 02_Clustering_KMeans.ipynb
│   └── results/02_clustering/
├── Unit_4_PBL2_HMM/                     Unit 4: Hidden Markov Model of daily demand regimes
│   ├── 03_HMM.ipynb
│   ├── 04_Comparison_and_Results.ipynb  (compares Units 3 and 4, final combined results)
│   └── results/03_hmm/, results/04_comparison/
├── Unit_5_PBL3_Gradient_Boosting/       Unit 5: gradient boosting, compared with PBL 1
│   ├── 06_PBL3_Gradient_Boosting.ipynb
│   └── results/06_boosting/
├── data/            raw Kaggle CSV (1,95,148 orders × 39 columns) — shared by all units
├── dataset/         cleaned and derived tables made by notebooks 01 and 05 — shared by all units
├── report/          final report (.docx, .pdf) and build/ = the scripts that generate it
├── presentation/    the three presentations (.pptx, .pdf) and their build scripts
├── requirements.txt          exact package versions
├── setup_venv_windows.bat    creates the virtual environment on Windows (double-click)
├── setup_venv_mac_linux.sh   creates the virtual environment on macOS / Linux
└── run_all_notebooks.py      runs all six notebooks in order and saves their outputs
```
Each unit's `results/` folder holds every chart (PNG), table (CSV), trained model (joblib) and summary (JSON) of
that unit. All notebooks already contain their outputs.

## 1. Create the virtual environment (once)
Install **Python 3.12** from python.org (hmmlearn has no Windows wheel for newer versions). Then:
- **Windows:** double-click `setup_venv_windows.bat`
- **macOS / Linux:** `bash setup_venv_mac_linux.sh`

This creates `.venv` in the project folder, installs the packages in `requirements.txt` and registers a Jupyter
kernel called **"Python (QuickCommerce ML .venv)"**. The environment is created on your own computer because a
virtual environment only works on the operating system and Python installation it was built with.

## 2. Run the notebooks
Run them in this order (later notebooks read the files written by earlier ones):

| Order | Unit folder | Notebook | Report chapter | What it does |
|---|---|---|---|---|
| 1 | Unit_1_2_PBL1_Data_and_Regression | `01_Create_Dataset.ipynb` | Ch. 3 (Unit 1) | cleaning, food filter, IPL-effect removal, store and daily tables |
| 2 | Unit_3_PBL2_Clustering_KMeans | `02_Clustering_KMeans.ipynb` | Ch. 4 (Unit 3) | feature selection, K-Means (k = 3), PCA, evaluation |
| 3 | Unit_4_PBL2_HMM | `03_HMM.ipynb` | Ch. 4 (Unit 4) | Gaussian HMM, Baum–Welch, Viterbi, validation, next-day forecast |
| 4 | Unit_4_PBL2_HMM | `04_Comparison_and_Results.ipynb` | Ch. 4 / 6 | same-data comparison, store cluster × regime |
| 5 | Unit_1_2_PBL1_Data_and_Regression | `05_PBL1_Regression.ipynb` | Ch. 3 (Units 1–2) | store-week dataset, Linear / Ridge / Bayesian regression |
| 6 | Unit_5_PBL3_Gradient_Boosting | `06_PBL3_Gradient_Boosting.ipynb` | Ch. 5 (Unit 5) | gradient boosting, time-series CV, comparison with PBL 1 |

- **All at once:** activate the environment (`.venv\Scripts\activate` on Windows, `source .venv/bin/activate`
  on macOS / Linux) and run `python run_all_notebooks.py` (about 2 minutes).
- **One by one (VS Code / Antigravity / Jupyter):** open a notebook, select the kernel
  "Python (QuickCommerce ML .venv)" and choose Run All.

Every notebook can be run from its unit folder; its first cell switches to the project folder so that `data/`,
`dataset/` and the other units' results are found. Running the notebooks regenerates every result exactly
(seed 42); only the measured training times differ slightly between computers.

**Readings in the report:** every number quoted in the report appears in a notebook output. Readings that are not
part of a table or chart above are printed in the last section of each notebook, **"Readings used in the report"**.

## Run on Google Colab
1. Upload the whole folder to Google Drive (e.g. `MyDrive/QuickCommerce_ML_Project_Final`).
2. Open a notebook in Colab and, in its first cell, remove the `#` from the three Colab lines
   (`pip install hmmlearn`, `drive.mount`, `%cd …`); adjust the `%cd` path to your folder.
3. Run all cells; run the notebooks in the order above.

## Web app
**Live app:** https://train-consist-management-app-ai-3xqjctuvkqvqztrz9p7euv.streamlit.app/

After creating the environment (step 1), start the dashboard with:

```
streamlit run app.py
```

The app only loads the saved models and results; it never retrains anything.

## Main results
| Unit | Result |
|---|---|
| 1–2 | Linear / Ridge / Bayesian regression: test MAE 5.35 orders per store-week (naive 6.76), R² 0.78 |
| 3 | K-Means: 3 store types (42 / 11 / 5), silhouette 0.42, stability ARI 0.85 |
| 4 | HMM: 3 regimes (Normal, Monsoon / disrupted, Festival peak), festival recall 82% |
| 5 | Gradient boosting (growth-free target): test MAE 5.32, R² 0.784 — best model |

## Rebuilding the report and presentations (optional)
- Report: `python report/build/build_report.py` (needs python-docx, LibreOffice and pandoc).
- Presentations: `node presentation/build_deck_pbl3.js` etc. (needs Node.js with pptxgenjs).

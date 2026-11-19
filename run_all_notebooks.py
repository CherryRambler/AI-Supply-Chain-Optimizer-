"""Run all six notebooks in order (01 -> 06), each from its own unit folder, saving the outputs into the notebooks.
Usage (with the virtual environment activated):  python run_all_notebooks.py"""
import os, subprocess, sys, time
ROOT = os.path.dirname(os.path.abspath(__file__))
ORDER = [("Unit_1_2_PBL1_Data_and_Regression", "01_Create_Dataset.ipynb"),
         ("Unit_3_PBL2_Clustering_KMeans", "02_Clustering_KMeans.ipynb"),
         ("Unit_4_PBL2_HMM", "03_HMM.ipynb"),
         ("Unit_4_PBL2_HMM", "04_Comparison_and_Results.ipynb"),
         ("Unit_1_2_PBL1_Data_and_Regression", "05_PBL1_Regression.ipynb"),
         ("Unit_5_PBL3_Gradient_Boosting", "06_PBL3_Gradient_Boosting.ipynb")]
for unit, nb in ORDER:
    t0 = time.time()
    print(f"running {unit}/{nb} ...", flush=True)
    subprocess.run([sys.executable, "-m", "jupyter", "nbconvert", "--to", "notebook", "--execute", "--inplace",
                    "--ExecutePreprocessor.timeout=1800", nb], cwd=os.path.join(ROOT, unit), check=True)
    print(f"   done in {time.time() - t0:.0f} s", flush=True)
print("All notebooks ran successfully.")

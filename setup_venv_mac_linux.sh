#!/usr/bin/env bash
# Creates the virtual environment .venv in this folder and installs every package the project needs.
# Run:  bash setup_venv_mac_linux.sh
set -e
cd "$(dirname "$0")"
PY=$(command -v python3.12 || command -v python3)
"$PY" -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m ipykernel install --user --name quickcommerce-ml --display-name "Python (QuickCommerce ML .venv)"
echo
echo "Done. The environment is in .venv"
echo "  Activate it:        source .venv/bin/activate"
echo "  Run all notebooks:  python run_all_notebooks.py"
echo "  Or open the notebooks in VS Code / Jupyter and choose the kernel \"Python (QuickCommerce ML .venv)\"."

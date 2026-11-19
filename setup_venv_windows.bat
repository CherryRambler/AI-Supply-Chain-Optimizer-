@echo off
REM Creates the virtual environment .venv in this folder and installs every package the project needs.
REM Double-click this file, or run it from a Command Prompt opened in the project folder.
cd /d "%~dp0"
py -3.12 -m venv .venv 2>nul || python -m venv .venv
if not exist .venv\Scripts\python.exe (
    echo Could not create the virtual environment. Install Python 3.12 from python.org and run this file again.
    pause & exit /b 1
)
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt || (echo Package installation failed. & pause & exit /b 1)
.venv\Scripts\python.exe -m ipykernel install --user --name quickcommerce-ml --display-name "Python (QuickCommerce ML .venv)"
echo.
echo Done. The environment is in .venv
echo   Activate it:        .venv\Scripts\activate
echo   Run all notebooks:  python run_all_notebooks.py
echo   Or open the notebooks in VS Code / Jupyter and choose the kernel "Python (QuickCommerce ML .venv)".
pause

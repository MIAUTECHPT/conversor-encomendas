@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Ambiente virtual nao encontrado nesta pasta!
    pause
    exit
)
call .venv\Scripts\activate
streamlit run app.py
pause
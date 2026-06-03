@echo off
cd /d %~dp0
python -m streamlit run streamlit_vimc_openmetadata_app.py
pause

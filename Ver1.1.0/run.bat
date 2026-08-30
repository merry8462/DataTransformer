@echo off
chcp 65001 >nul
cd /d "%~dp0"
python sql_excel_converter.py

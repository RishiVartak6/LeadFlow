@echo off
cd /d "%~dp0"
title LeadFlow

if exist "C:\Users\Rish\venv\Scripts\python.exe" (
    "C:\Users\Rish\venv\Scripts\python.exe" LeadFlow.py
) else (
    "C:\Program Files\Python310\python.exe" LeadFlow.py
)

pause

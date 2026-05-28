@echo off
title AutoReport Pro Cloud Edition
cd /d "%~dp0"
python -m pip install -r requirements.txt
python app.py
pause

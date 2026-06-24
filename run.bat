@echo off
rem Starts PasteTool in the background with no console window.
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0main.py"

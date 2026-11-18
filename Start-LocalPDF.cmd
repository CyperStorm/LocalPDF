@echo off
cd /d "%~dp0"
where pyw.exe >nul 2>&1
if errorlevel 1 goto python_fallback
start "" pyw.exe "%~dp0LocalPDF.py"
exit /b
:python_fallback
where pythonw.exe >nul 2>&1
if errorlevel 1 goto no_python
start "" pythonw.exe "%~dp0LocalPDF.py"
exit /b
:no_python
echo Install Python for Windows with Tcl/Tk and add it to PATH.
pause

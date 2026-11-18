@echo off
cd /d "%~dp0"
where py.exe >nul 2>&1
if errorlevel 1 goto python_fallback
py -m pip install --user Pillow PyMuPDF python-docx openpyxl python-pptx
goto done
:python_fallback
python -m pip install --user Pillow PyMuPDF python-docx openpyxl python-pptx
:done
if errorlevel 1 echo Setup failed. Check your Python installation and Internet connection.
pause


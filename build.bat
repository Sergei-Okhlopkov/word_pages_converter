@echo off
setlocal

set "VENV_DIR=.venv_build"

if not exist %VENV_DIR% (
    py -3 -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate
python -m pip install -r requirements.txt

pyinstaller --onefile --windowed --name WordWorkTool app.py

echo.
echo Build finished. EXE: dist\WordWorkTool.exe
endlocal

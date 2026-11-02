@echo off
cd /d "%~dp0"
echo Installing openpyxl and pyinstaller...
python -m pip install -q openpyxl pyinstaller
echo Building TreeCaseCompare.exe ...
python -m PyInstaller --noconfirm --onefile --name TreeCaseCompare compare_tree_cases.py
echo.
echo Done. Copy this file to the share drive:
echo   %cd%\dist\TreeCaseCompare.exe
echo Other PCs do NOT need Python. Just double-click the exe.
pause

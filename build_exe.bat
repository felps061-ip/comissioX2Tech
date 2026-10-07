@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=C:\Users\imper\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHONPATH=%ROOT%.vendor"

"%PYTHON%" -m PyInstaller --noconfirm --clean --onefile --windowed --name BEVI_2TECH_C6 --distpath "%ROOT%dist" --workpath "%ROOT%build" --specpath "%ROOT%build" "%ROOT%app_web.py"

echo.
echo Executavel gerado em:
echo %ROOT%dist\BEVI_2TECH_C6.exe
pause

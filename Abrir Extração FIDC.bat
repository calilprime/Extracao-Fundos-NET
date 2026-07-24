@echo off
REM ===========================================================
REM  Extracao FIDC - FNET (Netz Asset)
REM  Da um duplo-clique neste arquivo para abrir a interface.
REM ===========================================================
cd /d "%~dp0"

REM Tenta 'py' (launcher do Windows) e depois 'python'
where py >nul 2>nul
if %errorlevel%==0 (
    py fnet_app.py
) else (
    python fnet_app.py
)

if %errorlevel% neq 0 (
    echo.
    echo Ocorreu um erro. Verifique se o Python esta instalado e se rodou:
    echo     pip install openpyxl
    echo.
    pause
)

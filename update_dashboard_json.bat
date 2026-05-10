@echo off
setlocal

cd /d "%~dp0"

echo ============================================================
echo  Dashboard data update
echo ============================================================
echo.
echo Source:
echo   Data Update OpenAPI URLs.md
echo.
echo The script will print every tab/data item result below.
echo This window will stay open until you press a key.
echo.

set "SCRIPT=update_openapi_dashboard.py"
set "PYTHON_EXE=%USERPROFILE%\anaconda3\python.exe"

if not exist "%SCRIPT%" (
  echo [FAILED] Cannot find %SCRIPT%.
  echo Please check that this bat file is in the same folder as index.html.
  echo.
  echo Press any key to close this window.
  pause >nul
  exit /b 1
)

if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" "%SCRIPT%"
) else (
  python "%SCRIPT%"
)

set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" (
  echo [FAILED] Update did not finish. Please check the messages above.
) else (
  echo [DONE] Update check finished.
)
echo.
echo Press any key to close this window after reviewing all results.
pause >nul
exit /b %RESULT%

@echo off
REM manage_assets.bat - Windows wrapper for CityInfraXLS

python manage_assets.py %*
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo Command failed with error code %ERRORLEVEL%
  pause
)
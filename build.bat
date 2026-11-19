@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"
python build.py %*
if errorlevel 1 (
  echo BUILD FAILED
  exit /b 1
)
echo Done. See dist\ for the exe.

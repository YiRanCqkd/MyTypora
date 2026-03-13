@echo off
setlocal

chcp 65001 >nul
cd /d "%~dp0"

set "TARGET_EXE=%~1"
if "%TARGET_EXE%"=="" set "TARGET_EXE=%CD%\release\electron\win-unpacked\UniReadMD.exe"

set "ICON_FILE=%CD%\res\unireadmd.ico"
set "RCEDIT_EXE=%CD%\node_modules\electron-winstaller\vendor\rcedit.exe"

if not exist "%RCEDIT_EXE%" (
  echo [ERROR] rcedit not found: "%RCEDIT_EXE%"
  echo         Please run: npm install
  exit /b 1
)

if not exist "%ICON_FILE%" (
  echo [ERROR] icon not found: "%ICON_FILE%"
  exit /b 1
)

if not exist "%TARGET_EXE%" (
  echo [ERROR] exe not found: "%TARGET_EXE%"
  echo.
  echo Usage:
  echo   patch_exe_icon.bat "I:\path\to\UniReadMD.exe"
  exit /b 1
)

"%RCEDIT_EXE%" "%TARGET_EXE%" --set-icon "%ICON_FILE%" ^
  --set-version-string ProductName UniReadMD ^
  --set-version-string FileDescription UniReadMD ^
  --set-version-string CompanyName UniReadMD ^
  --set-version-string OriginalFilename UniReadMD.exe

if errorlevel 1 (
  echo [ERROR] patch failed.
  exit /b 1
)

echo [OK] icon/resource patched:
echo      "%TARGET_EXE%"
exit /b 0

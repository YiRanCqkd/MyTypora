@echo off
setlocal

chcp 65001 >nul
cd /d "%~dp0"

set "TARGET_EXE=%~1"
if "%TARGET_EXE%"=="" (
  if exist "%CD%\release\electron\UniReadMD.exe" (
    set "TARGET_EXE=%CD%\release\electron\UniReadMD.exe"
  ) else (
    set "TARGET_EXE=%CD%\release\electron\win-unpacked\UniReadMD.exe"
  )
)

if not exist "%TARGET_EXE%" (
  echo [ERROR] Executable not found:
  echo         "%TARGET_EXE%"
  echo.
  echo Usage:
  echo   start_from_exe.bat "I:\path\to\UniReadMD.exe" [extra args]
  exit /b 1
)

shift
start "" "%TARGET_EXE%" %*

exit /b 0

@echo off
setlocal

chcp 65001 >nul
cd /d "%~dp0"

echo 启动开发环境...
call npm run dev
if errorlevel 1 goto :error

goto :eof

:error
echo.
echo [ERROR] 启动失败，请根据上方日志排查。
exit /b 1

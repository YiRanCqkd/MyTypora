@echo off
setlocal

echo Building MyTypora Markdown Editor...
echo.

REM Set UTF-8 code page for readable output
chcp 65001 >nul 2>nul

set "VSWHERE=C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
set "MSBUILD="

REM 1) Prefer vswhere if present
if exist "%VSWHERE%" (
    for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath`) do (
        if exist "%%i\MSBuild\Current\Bin\MSBuild.exe" (
            set "MSBUILD=%%i\MSBuild\Current\Bin\MSBuild.exe"
        )
    )
)

REM 2) Fallbacks (common install locations)
if not defined MSBUILD if exist "D:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=D:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
if not defined MSBUILD if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
if not defined MSBUILD if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
if not defined MSBUILD if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
if not defined MSBUILD if exist "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"

if not defined MSBUILD (
    echo.
    echo ERROR: Visual Studio 2022 MSBuild not found.
    echo Checked vswhere and common install paths.
    echo.
    echo Tips:
    echo 1. Install Visual Studio 2022 or Build Tools
    echo 2. Ensure the "MSBuild" component is installed
    echo.
    pause
    exit /b 1
)

echo Using MSBuild: "%MSBUILD%"
echo.

REM Phase 1 target: x86 Release only
echo Building x86 Release...
"%MSBUILD%" MarkdownEditor.sln /t:Build /p:Configuration=Release /p:Platform=x86 /v:minimal
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Build failed.
    echo Error code: %ERRORLEVEL%
    echo.
    echo Common solutions:
    echo 1. Ensure "Desktop development with C++" is installed
    echo 2. Ensure MFC x86 components are installed
    echo 3. Open MarkdownEditor.sln in Visual Studio and build x86 Release
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Build successful.
echo Output: Release\MarkdownEditor.exe
echo.
pause

endlocal

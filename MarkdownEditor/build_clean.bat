@echo off
setlocal

echo Cleaning and rebuilding MyTypora Markdown Editor...
echo.

chcp 65001 >nul 2>nul

set "VSWHERE=C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
set "MSBUILD="

if exist "%VSWHERE%" (
    for /f "usebackq tokens=*" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath`) do (
        if exist "%%i\MSBuild\Current\Bin\MSBuild.exe" (
            set "MSBUILD=%%i\MSBuild\Current\Bin\MSBuild.exe"
        )
    )
)

if not defined MSBUILD if exist "D:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" set "MSBUILD=D:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"

if not defined MSBUILD (
    echo ERROR: MSBuild not found.
    pause
    exit /b 1
)

echo Using MSBuild: "%MSBUILD%"
echo.

echo Cleaning x86 Release...
"%MSBUILD%" MarkdownEditor.sln /t:Clean /p:Configuration=Release /p:Platform=x86 /v:minimal

echo.
echo Rebuilding x86 Release...
"%MSBUILD%" MarkdownEditor.sln /t:Rebuild /p:Configuration=Release /p:Platform=x86 /v:minimal

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Rebuild failed with error code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo Rebuild successful.
echo Output: Release\MarkdownEditor.exe
echo.
pause

endlocal

@echo off
setlocal

rem Use the solution directory even when launched from another working directory.
pushd "%~dp0"
if errorlevel 1 (
    echo Could not open the solution directory.
    pause
    exit /b 1
)

dotnet publish "Msg2EmlConverter\Msg2EmlConverter.csproj" ^
    -c Release ^
    -r win-x64 ^
    --self-contained true ^
    -p:PublishSingleFile=true ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -o "publish"

set "publishExitCode=%errorlevel%"
if "%publishExitCode%"=="0" (
    echo Published successfully: "%~dp0publish\Msg2EmlConverter.exe"
) else (
    echo Publishing failed with exit code %publishExitCode%.
)

popd
rem Keep the result visible when started by double-clicking the script.
pause
exit /b %publishExitCode%

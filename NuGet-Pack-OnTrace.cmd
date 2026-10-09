@echo off
REM Builds the OnTrace NuGet package into .\artifacts
setlocal
set OUT=%~dp0artifacts
dotnet pack OnTrace\OnTrace.vbproj -c Release -o "%OUT%"
echo.
echo Package written to: %OUT%
endlocal

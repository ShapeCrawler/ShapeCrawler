@echo off
setlocal

set "projectPath=..\src\ShapeCrawler\ShapeCrawler.csproj"
set "outputDirectory=.\Output"
set "zipExtension=.zip"

REM Pack the project into a NuGet package using Release configuration
powershell -Command "& { dotnet pack '%projectPath%' --configuration Release -o '%outputDirectory%' }"

if %errorlevel% neq 0 (
    echo FAILED
    exit /b %errorlevel%
)

REM Get the name of the generated .nupkg file
for /f "delims=" %%i in ('powershell -Command "& { (Get-ChildItem -Path '%outputDirectory%\*.nupkg' | Select-Object -First 1).FullName }"') do set "nupkgFile=%%i"

REM Check if .nupkg file was found
if not defined nupkgFile (
    echo No .nupkg file found
    exit /b 1
)

REM Create a ZIP file from the .nupkg
powershell -Command "& { Compress-Archive -Path '%nupkgFile%' -DestinationPath '%nupkgFile%%zipExtension%' }"

if %errorlevel% eq 0 (
    echo SUCCESS
    color 2A
) else (
    echo FAILED
    color 4C
)

echo Press any key to continue ...
pause > nul

endlocal
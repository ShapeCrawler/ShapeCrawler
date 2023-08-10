# Pack the project into a NuGet package using Release configuration
dotnet pack "..\src\ShapeCrawler\ShapeCrawler.csproj" --configuration Release -o .\Output

# Get the name of the generated .nupkg file
$nupkgFile = Get-ChildItem -Path .\Output\*.nupkg | Select-Object -First 1

# Create a ZIP file from the .nupkg
Compress-Archive -Path $nupkgFile.FullName -DestinationPath "$($nupkgFile.FullName).zip"

Write-Host "Press any key to continue ..."
$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") > $null

param(
    [string]$Version = "2.1.14"
)

$ErrorActionPreference = "Stop"

$pyArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--onefile",
    "--windowed",
    "--name", "ShotlistCreator",
    "--add-data", "assets;assets"
)

if (Test-Path "icon.ico") {
    $pyArgs += @("--icon", "icon.ico")
}
elseif (Test-Path "icon.png") {
    $pyArgs += @("--icon", "icon.png")
}

$pyArgs += "ShotlistCreator.py"

python @pyArgs

$sourceExe = "dist/ShotlistCreator.exe"
$targetExe = "dist/ShotlistCreator-$Version-Windows-Setup.exe"

if (-Not (Test-Path $sourceExe)) {
    throw "PyInstaller output not found: $sourceExe"
}

Copy-Item -Path $sourceExe -Destination $targetExe -Force
Write-Host "Built: $targetExe"

param(
    [string]$Version = "2.1.14"
)

$ErrorActionPreference = "Stop"

$iconArg = ""
if (Test-Path "icon.ico") {
    $iconArg = "--icon icon.ico"
}
elseif (Test-Path "icon.png") {
    $iconArg = "--icon icon.png"
}

python -m PyInstaller --noconfirm --windowed --name ShotlistCreator --add-data "assets;assets" $iconArg ShotlistCreator.py

$issPath = "packaging/windows/ShotlistCreator.iss"
$innoExe = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if (-Not (Test-Path $innoExe)) {
    throw "Inno Setup compiler not found: $innoExe"
}

& $innoExe "/DMyAppVersion=$Version" $issPath
Write-Host "Built: dist/ShotlistCreator-$Version-Windows-Setup.exe"

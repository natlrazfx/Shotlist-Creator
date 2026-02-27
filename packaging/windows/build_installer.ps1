param(
    [string]$Version = "2.1.15"
)

$ErrorActionPreference = "Stop"

$pyArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
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

$issPath = "packaging/windows/ShotlistCreator.iss"
$innoExe = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if (-Not (Test-Path $innoExe)) {
    throw "Inno Setup compiler not found: $innoExe"
}

& $innoExe "/DMyAppVersion=$Version" $issPath
Write-Host "Built: dist/ShotlistCreator-$Version-Windows-Setup.exe"

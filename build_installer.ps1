$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Write-Host "Installing build dependencies..."
py -m pip install --upgrade pip
py -m pip install -r requirements.txt

Write-Host "Building standalone application..."
py -m PyInstaller --noconfirm --clean CorpusQuerier.spec

$PossibleCompilers = @(
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
)
$InnoCompiler = $PossibleCompilers | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $InnoCompiler) {
    throw "Inno Setup 6 was not found. Install it from https://jrsoftware.org/isdl.php and run this script again."
}

Write-Host "Building Windows installer..."
& $InnoCompiler "installer\CorpusQuerier.iss"

Write-Host ""
Write-Host "Installer created successfully:"
Write-Host "$ProjectRoot\installer-output\Corpus-Querier-Setup.exe"


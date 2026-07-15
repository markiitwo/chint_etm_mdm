$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$AppName = "CHINT ETM MDM"
$VenvDir = Join-Path $Root ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$DistDir = Join-Path $Root "dist"
$AppDir = Join-Path $DistDir $AppName
$ZipPath = Join-Path $DistDir "CHINT_ETM_MDM_portable.zip"
$BundledRules = "src\chint_etm_mdm\rules\attribute_mappings.json;chint_etm_mdm\rules"

function New-LocalVenv {
    if (Test-Path $VenvPython) {
        Write-Host "Local venv already exists: $VenvDir"
        return
    }

    Write-Host "Creating local Python venv in .venv..."

    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $VenvDir
        Assert-LastExitCode "venv creation"
        return
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $VenvDir
        Assert-LastExitCode "venv creation"
        return
    }

    throw "Python was not found. Install Python 3.10+ to build the portable app."
}

function Assert-LastExitCode {
    param([string]$StepName)
    if ($LASTEXITCODE -ne 0) {
        throw "$StepName failed with exit code $LASTEXITCODE."
    }
}

New-LocalVenv

if (-not (Test-Path $VenvPython)) {
    throw "Python was not found inside .venv: $VenvPython"
}

Write-Host "Updating build tools..."
& $VenvPython -m pip install --upgrade pip
Assert-LastExitCode "pip upgrade"
& $VenvPython -m pip install -r requirements.txt pyinstaller
Assert-LastExitCode "dependency install"

Write-Host "Building portable app..."
& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name $AppName `
    --paths src `
    --add-data $BundledRules `
    run_app.py
Assert-LastExitCode "PyInstaller"

$ExePath = Join-Path $AppDir "$AppName.exe"
if (-not (Test-Path $ExePath)) {
    throw "Build finished without the expected EXE: $ExePath"
}

if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

Write-Host "Creating portable zip..."
Compress-Archive -Path $AppDir -DestinationPath $ZipPath -Force

Write-Host ""
Write-Host "Done."
Write-Host "App folder: $AppDir"
Write-Host "Portable zip: $ZipPath"

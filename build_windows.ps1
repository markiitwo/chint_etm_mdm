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
        Write-Host "Локальная среда уже есть: $VenvDir"
        return
    }

    Write-Host "Создаю локальную Python-среду в .venv..."

    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $VenvDir
        return
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        & python -m venv $VenvDir
        return
    }

    throw "Python не найден. Установите Python 3.10+ для сборки portable-версии."
}

New-LocalVenv

if (-not (Test-Path $VenvPython)) {
    throw "Не найден Python внутри .venv: $VenvPython"
}

Write-Host "Обновляю инструменты сборки..."
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r requirements.txt pyinstaller

Write-Host "Собираю portable-приложение..."
& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onedir `
    --name $AppName `
    --paths src `
    --add-data $BundledRules `
    run_app.py

$ExePath = Join-Path $AppDir "$AppName.exe"
if (-not (Test-Path $ExePath)) {
    throw "Сборка завершилась без ожидаемого EXE: $ExePath"
}

if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

Write-Host "Упаковываю portable-архив..."
Compress-Archive -Path $AppDir -DestinationPath $ZipPath -Force

Write-Host ""
Write-Host "Готово."
Write-Host "Папка программы: $AppDir"
Write-Host "Архив для передачи: $ZipPath"

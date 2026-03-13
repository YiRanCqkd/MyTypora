$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$stageOutputName = 'electron-stage-' + [guid]::NewGuid().ToString('N')
$prepackagedName = 'electron-prepackaged-' + [guid]::NewGuid().ToString('N')
$installerOutputName = 'electron-installer-' + [guid]::NewGuid().ToString('N')
$stageOutputDir = Join-Path $projectRoot ('release\' + $stageOutputName)
$prepackagedDir = Join-Path $projectRoot ('release\' + $prepackagedName)
$installerOutputDir = Join-Path $projectRoot ('release\' + $installerOutputName)
$targetDir = Join-Path $projectRoot 'release\electron'
$targetInstaller = Join-Path $targetDir 'UniReadMD-Setup.exe'
$legacyPortableExe = Join-Path $targetDir 'UniReadMD.exe'
$legacyUnpackedDir = Join-Path $targetDir 'win-unpacked'
$targetIcon = Join-Path $targetDir 'UniReadMD.ico'
$builder = Join-Path $projectRoot 'node_modules\.bin\electron-builder.cmd'
$rcedit = Join-Path $projectRoot 'node_modules\electron-winstaller\vendor\rcedit.exe'
$sourceIcon = Join-Path $projectRoot 'res\unireadmd.ico'

if (-not (Test-Path $builder)) {
  throw 'electron-builder command not found.'
}
if (-not (Test-Path $rcedit)) {
  throw 'rcedit executable not found.'
}

Push-Location $projectRoot
try {
  & $builder --config 'electron\builder.json' --win --dir `
    "--config.directories.output=release/$stageOutputName"
  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $stageAppDir = Join-Path $stageOutputDir 'win-unpacked'
  $patchedExe = Join-Path $prepackagedDir 'UniReadMD.exe'
  if (-not (Test-Path $stageAppDir)) {
    throw 'Unpacked app directory not found after dir packaging.'
  }

  Copy-Item $stageAppDir $prepackagedDir -Recurse -Force
  Start-Sleep -Milliseconds 800

  $patched = $false
  foreach ($attempt in 1..5) {
    & $rcedit $patchedExe `
      --set-icon $sourceIcon `
      --set-version-string ProductName UniReadMD `
      --set-version-string FileDescription UniReadMD `
      --set-version-string CompanyName UniReadMD `
      --set-version-string OriginalFilename UniReadMD.exe
    if ($LASTEXITCODE -eq 0) {
      $patched = $true
      break
    }

    Start-Sleep -Milliseconds (600 * $attempt)
  }

  if (-not $patched) {
    throw 'Failed to patch icon for prepackaged app executable.'
  }

  & $builder --config 'electron\builder.json' --win nsis `
    --prepackaged "release/$prepackagedName" `
    "--config.directories.output=release/$installerOutputName"
  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $sourceInstaller = Get-ChildItem $installerOutputDir -Filter '*.exe' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if (-not $sourceInstaller) {
    throw 'Installer exe not found after nsis packaging.'
  }

  New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
  if (Test-Path $legacyPortableExe) {
    Remove-Item $legacyPortableExe -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $legacyUnpackedDir) {
    Remove-Item $legacyUnpackedDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  Copy-Item $sourceInstaller.FullName $targetInstaller -Force
  if (Test-Path $sourceIcon) {
    Copy-Item $sourceIcon $targetIcon -Force
  }

  Write-Output "[OK] installer synced: $targetInstaller"
} finally {
  Pop-Location

  if (Test-Path $stageOutputDir) {
    Remove-Item $stageOutputDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $prepackagedDir) {
    Remove-Item $prepackagedDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $installerOutputDir) {
    Remove-Item $installerOutputDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}

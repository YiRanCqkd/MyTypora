$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$stageOutputName = 'electron-stage-' + [guid]::NewGuid().ToString('N')
$portableOutputName = 'electron-portable-' + [guid]::NewGuid().ToString('N')
$stageOutputDir = Join-Path $projectRoot ('release\' + $stageOutputName)
$portableOutputDir = Join-Path $projectRoot ('release\' + $portableOutputName)
$targetDir = Join-Path $projectRoot 'release\electron'
$targetExe = Join-Path $targetDir 'UniReadMD.exe'
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
  & $builder --config 'electron\builder.json' --win --dir "--config.directories.output=release/$stageOutputName"
  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $prepackagedDir = Join-Path $stageOutputDir 'win-unpacked'
  $innerExe = Join-Path $prepackagedDir 'UniReadMD.exe'
  if (-not (Test-Path $innerExe)) {
    throw 'Unpacked app executable not found after dir packaging.'
  }

  & $rcedit $innerExe `
    --set-icon $sourceIcon `
    --set-version-string ProductName UniReadMD `
    --set-version-string FileDescription UniReadMD `
    --set-version-string CompanyName UniReadMD `
    --set-version-string OriginalFilename UniReadMD.exe
  if ($LASTEXITCODE -ne 0) {
    throw 'Failed to patch icon for unpacked app executable.'
  }

  & $builder --config 'electron\builder.json' --win portable `
    --prepackaged "release/$stageOutputName/win-unpacked" `
    "--config.directories.output=release/$portableOutputName"
  if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
  }

  $sourceExe = Join-Path $portableOutputDir 'UniReadMD.exe'
  if (-not (Test-Path $sourceExe)) {
    throw 'Portable exe not found after portable packaging.'
  }

  New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
  Copy-Item $sourceExe $targetExe -Force
  if (Test-Path $sourceIcon) {
    Copy-Item $sourceIcon $targetIcon -Force
  }

  Write-Output "[OK] portable exe synced: $targetExe"
} finally {
  Pop-Location

  if (Test-Path $stageOutputDir) {
    Remove-Item $stageOutputDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  if (Test-Path $portableOutputDir) {
    Remove-Item $portableOutputDir -Recurse -Force -ErrorAction SilentlyContinue
  }
}

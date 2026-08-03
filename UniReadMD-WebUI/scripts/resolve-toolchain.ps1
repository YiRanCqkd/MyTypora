[CmdletBinding()]
param(
  [switch]$AsObject,
  [string]$BaselinePath = (Join-Path $PSScriptRoot '..\release\toolchain-baseline.json'),
  [string[]]$PathEntriesOverride
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ExpectedNodePackageId = 'OpenJS.NodeJS'
$ExpectedNodePackageVersion = '24.8.0'
$ExpectedNodeVersion = 'v24.8.0'
$ExpectedGitPackageId = 'Git.Git'
$ExpectedGitPackageVersion = '2.47.1.2'
$ExpectedGitVersion = 'git version 2.47.1.windows.2'

function Exit-WithCode {
  param(
    [string]$Code,
    [int]$ExitCode,
    [string]$Detail = ''
  )

  if ($Detail.Length -gt 0) {
    [Console]::Error.WriteLine("${Code}: ${Detail}")
  } else {
    [Console]::Error.WriteLine($Code)
  }

  exit $ExitCode
}

function Get-UninstallEntries {
  $roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  @(Get-ItemProperty -Path $roots -ErrorAction SilentlyContinue)
}

function ConvertTo-RegisteredPathEntry {
  param([string]$Entry)

  $trimmed = $Entry.Trim().Trim('"')
  if ([string]::IsNullOrWhiteSpace($trimmed)) {
    return $null
  }

  $expanded = [Environment]::ExpandEnvironmentVariables($trimmed)
  try {
    $fullPath = [IO.Path]::GetFullPath($expanded)
  } catch {
    return $null
  }

  $resolved = Resolve-Path -LiteralPath $fullPath -ErrorAction SilentlyContinue
  if ($null -eq $resolved) {
    return $null
  }

  $resolved.ProviderPath.TrimEnd('\')
}

function Get-RegisteredPathEntries {
  $entries = New-Object 'System.Collections.Generic.List[string]'

  foreach ($scope in @('Machine', 'User')) {
    $rawValue = [Environment]::GetEnvironmentVariable('Path', $scope)
    if ([string]::IsNullOrWhiteSpace($rawValue)) {
      continue
    }

    foreach ($entry in $rawValue.Split(';')) {
      $normal = ConvertTo-RegisteredPathEntry $entry
      if ($null -eq $normal) {
        continue
      }

      if (-not $entries.Contains($normal)) {
        $entries.Add($normal)
      }
    }
  }

  @($entries)
}

function Get-PathEntries {
  param([string[]]$Override)

  if ($PSBoundParameters.ContainsKey('Override') -and $null -ne $Override) {
    $entries = @()
    foreach ($entry in $Override) {
      $normal = ConvertTo-RegisteredPathEntry $entry
      if ($null -ne $normal) {
        $entries += $normal
      }
    }

    return @($entries)
  }

  Get-RegisteredPathEntries
}

function Invoke-Version {
  param(
    [string]$FilePath,
    [string[]]$Arguments
  )

  $output = & $FilePath @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "version command failed: $FilePath"
  }

  $firstLine = @($output | Select-Object -First 1)[0]
  if ($null -eq $firstLine) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "version command returned no output: $FilePath"
  }

  $firstLine.ToString().Trim()
}

function Get-HashLower {
  param([string]$Path)

  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "required file missing: $Path"
  }

  (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-SinglePackage {
  param(
    [object[]]$Entries,
    [string]$DisplayName,
    [string]$DisplayVersion
  )

  $matches = @(
    $Entries | Where-Object {
      $name = $_.PSObject.Properties['DisplayName']
      $version = $_.PSObject.Properties['DisplayVersion']
      $null -ne $name -and
        $null -ne $version -and
        $name.Value -eq $DisplayName -and
        $version.Value -eq $DisplayVersion
    }
  )

  if ($matches.Count -ne 1) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 (
      "package candidate count for ${DisplayName} ${DisplayVersion}: $($matches.Count)"
    )
  }

  $matches[0]
}

function Resolve-NodeToolchain {
  param([string[]]$PathEntries)

  $candidates = @()
  foreach ($entry in $PathEntries) {
    $nodePath = Join-Path $entry 'node.exe'
    $npmCli = Join-Path $entry 'node_modules\npm\bin\npm-cli.js'
    if (-not (Test-Path -LiteralPath $nodePath -PathType Leaf)) {
      continue
    }

    $nodeVersion = Invoke-Version $nodePath @('--version')
    if ($nodeVersion -ne $ExpectedNodeVersion) {
      continue
    }

    if (-not (Test-Path -LiteralPath $npmCli -PathType Leaf)) {
      Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "npm cli missing beside node: $entry"
    }

    $npmVersion = Invoke-Version $nodePath @($npmCli, '--version')
    $candidates += [pscustomobject]@{
      Node = (Resolve-Path -LiteralPath $nodePath).ProviderPath
      Root = (Resolve-Path -LiteralPath $entry).ProviderPath.TrimEnd('\')
      NpmCli = (Resolve-Path -LiteralPath $npmCli).ProviderPath
      Version = $nodeVersion
      NpmVersion = $npmVersion
      Sha256 = Get-HashLower $nodePath
      NpmSha256 = Get-HashLower $npmCli
    }
  }

  if ($candidates.Count -ne 1) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 (
      "node $ExpectedNodeVersion candidate count: $($candidates.Count)"
    )
  }

  $candidates[0]
}

function Resolve-GitToolchain {
  param(
    [object]$GitPackage,
    [string[]]$PathEntries
  )

  $installRoot = $GitPackage.InstallLocation
  if ([string]::IsNullOrWhiteSpace($installRoot)) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 'Git install location missing'
  }

  $gitPath = Join-Path $installRoot 'cmd\git.exe'
  if (-not (Test-Path -LiteralPath $gitPath -PathType Leaf)) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "Git executable missing: $gitPath"
  }

  $resolvedGit = (Resolve-Path -LiteralPath $gitPath).ProviderPath
  $gitParent = (Split-Path -Parent $resolvedGit).TrimEnd('\')
  $registered = @($PathEntries | Where-Object {
    [StringComparer]::OrdinalIgnoreCase.Equals($_.TrimEnd('\'), $gitParent)
  })

  if ($registered.Count -ne 1) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 (
      "Git PATH registration count for ${gitParent}: $($registered.Count)"
    )
  }

  $version = Invoke-Version $resolvedGit @('--version')
  if ($version -ne $ExpectedGitVersion) {
    Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "Git version mismatch: $version"
  }

  [pscustomobject]@{
    Git = $resolvedGit
    Version = $version
    Sha256 = Get-HashLower $resolvedGit
  }
}

function Assert-Equal {
  param(
    [string]$Name,
    [string]$Expected,
    [string]$Actual
  )

  if ($Expected -cne $Actual) {
    Exit-WithCode 'ERR_TOOLCHAIN_BASELINE_CHANGED' 71 "$Name expected '$Expected' got '$Actual'"
  }
}

function Assert-PathEqual {
  param(
    [string]$Name,
    [string]$Expected,
    [string]$Actual
  )

  if (-not [StringComparer]::OrdinalIgnoreCase.Equals($Expected, $Actual)) {
    Exit-WithCode 'ERR_TOOLCHAIN_BASELINE_CHANGED' 71 "$Name expected '$Expected' got '$Actual'"
  }
}

function Assert-Baseline {
  param(
    [object]$Baseline,
    [object]$Node,
    [object]$Git
  )

  Assert-Equal 'node.packageId' $ExpectedNodePackageId $Baseline.node.packageId
  Assert-Equal 'node.packageVersion' $ExpectedNodePackageVersion $Baseline.node.packageVersion
  Assert-Equal 'node.version' $Node.Version $Baseline.node.version
  Assert-PathEqual 'node.path' $Node.Node $Baseline.node.path
  Assert-Equal 'node.sha256' $Node.Sha256 $Baseline.node.sha256

  Assert-Equal 'npm.packageId' $ExpectedNodePackageId $Baseline.npm.packageId
  Assert-Equal 'npm.packageVersion' $ExpectedNodePackageVersion $Baseline.npm.packageVersion
  Assert-Equal 'npm.version' $Node.NpmVersion $Baseline.npm.version
  Assert-PathEqual 'npm.cliPath' $Node.NpmCli $Baseline.npm.cliPath
  Assert-PathEqual 'npm.nodeRoot' $Node.Root $Baseline.npm.nodeRoot
  Assert-Equal 'npm.sha256' $Node.NpmSha256 $Baseline.npm.sha256

  Assert-Equal 'git.packageId' $ExpectedGitPackageId $Baseline.git.packageId
  Assert-Equal 'git.packageVersion' $ExpectedGitPackageVersion $Baseline.git.packageVersion
  Assert-Equal 'git.version' $Git.Version $Baseline.git.version
  Assert-PathEqual 'git.path' $Git.Git $Baseline.git.path
  Assert-Equal 'git.sha256' $Git.Sha256 $Baseline.git.sha256
}

$resolvedBaselinePath = [IO.Path]::GetFullPath($BaselinePath)
if (-not (Test-Path -LiteralPath $resolvedBaselinePath -PathType Leaf)) {
  Exit-WithCode 'ERR_BOOTSTRAP_TOOLCHAIN' 70 "baseline missing: $resolvedBaselinePath"
}

$uninstallEntries = Get-UninstallEntries
$null = Get-SinglePackage $uninstallEntries 'Node.js' $ExpectedNodePackageVersion
$gitPackage = Get-SinglePackage $uninstallEntries 'Git' $ExpectedGitPackageVersion
$pathEntries = if ($PSBoundParameters.ContainsKey('PathEntriesOverride')) {
  Get-PathEntries $PathEntriesOverride
} else {
  Get-RegisteredPathEntries
}
$node = Resolve-NodeToolchain $pathEntries
$git = Resolve-GitToolchain $gitPackage $pathEntries
try {
  $baseline = Get-Content -LiteralPath $resolvedBaselinePath -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
  Exit-WithCode 'ERR_TOOLCHAIN_BASELINE_CHANGED' 71 'baseline is malformed JSON'
}

Assert-Baseline $baseline $node $git

$result = [pscustomobject]@{
  Node = $node.Node
  NpmCli = $node.NpmCli
  Git = $git.Git
  NodeVersion = $node.Version
  NpmVersion = $node.NpmVersion
  GitVersion = $git.Version
}

if ($AsObject) {
  $result
} else {
  $result | ConvertTo-Json -Depth 4
}

param(
  [string]$Tool
)

$ErrorActionPreference = 'Stop'

if ($Tool -notin @('npm', 'node')) {
  [Console]::Error.WriteLine('ERR_INVOKE_NODE_TOOL: tool 必须是 npm 或 node')
  exit 64
}

$RawArguments = @($args)
$LiteralMarkerIndexes = @(
  0..([Math]::Max(0, $RawArguments.Count - 1)) | Where-Object {
    $RawArguments[$_] -eq '-LiteralArguments'
  }
)
if ($LiteralMarkerIndexes.Count -gt 1 -or (
  $LiteralMarkerIndexes.Count -eq 1 -and $LiteralMarkerIndexes[0] -ne 0
)) {
  [Console]::Error.WriteLine(
    'ERR_INVOKE_NODE_TOOL_ARGUMENTS: 不得混用兼容参数与 -LiteralArguments'
  )
  exit 64
}

$UsesLiteralArguments = $LiteralMarkerIndexes.Count -eq 1
if ($UsesLiteralArguments -and $RawArguments.Count -gt 1) {
  $ToolArguments = @($RawArguments[1..($RawArguments.Count - 1)])
} elseif ($UsesLiteralArguments) {
  $ToolArguments = @()
} else {
  $ToolArguments = @($RawArguments)
}

# npm 脚本参数含独立 -- 时，调用方必须通过 -LiteralArguments 显式传入完整数组。

$BaselinePath = Join-Path $PSScriptRoot '..\release\toolchain-baseline.json'

try {
  $toolchain = & (Join-Path $PSScriptRoot 'resolve-toolchain.ps1') `
    -BaselinePath $BaselinePath `
    -AsObject

  if ($Tool -eq 'node') {
    & $toolchain.Node @ToolArguments
    exit $LASTEXITCODE
  }

  # npm 生命周期内部会执行 node；仅向子进程暴露已验证的固定 Node 安装目录。
  $previousPath = $env:PATH
  try {
    $systemDirectory = "$($env:SystemRoot)\System32"
    $powerShellDirectory = Join-Path $systemDirectory 'WindowsPowerShell\v1.0'
    $nodeDirectory = Split-Path -Parent $toolchain.Node
    $env:PATH = @(
      $nodeDirectory,
      $systemDirectory,
      $powerShellDirectory
    ) -join [IO.Path]::PathSeparator
    # Playwright 需要 taskkill.exe，electron-builder 需要系统 powershell.exe。
    & $toolchain.Node $toolchain.NpmCli @ToolArguments
    exit $LASTEXITCODE
  } finally {
    $env:PATH = $previousPath
  }
} catch {
  [Console]::Error.WriteLine("ERR_INVOKE_NODE_TOOL: $($_.Exception.Message)")
  exit 72
}

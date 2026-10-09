param(
    [string]$Chapter = "用户手册 > iFC用户手册",
    [string]$Hdx = "",
    [string]$Out = "",
    [switch]$DryRun,
    [switch]$Overwrite,
    [switch]$AllowLarge,
    [double]$MaxMB = 0,
    [string]$SplitDir = "",
    [string]$SplitPrefix = "",
    [switch]$AllowLargeAttachments,
    [switch]$MergeWordAttachments
)
$ErrorActionPreference = "Stop"
# Use the caller's Python installation, not a machine-specific Codex runtime.
$wordPython = Get-Command python -ErrorAction SilentlyContinue
if (-not $wordPython) { $wordPython = Get-Command py -ErrorAction SilentlyContinue }
if (-not $wordPython) {
    Write-Error "未找到 Python。请先安装 Python 3.10 或更新版本。"
    exit 2
}
$wordProjectRoot = Split-Path -Parent $PSScriptRoot
Push-Location -LiteralPath $wordProjectRoot
try {
    $wordArguments = @((Join-Path $PSScriptRoot "export_word.py"), "--chapter", $Chapter)
    if ($Hdx) { $wordArguments += @("--hdx", $Hdx) }
    if ($Out) { $wordArguments += @("--out", $Out) }
    if ($DryRun) { $wordArguments += "--dry-run" }
    if ($Overwrite) { $wordArguments += "--overwrite" }
    if ($AllowLarge) { $wordArguments += "--allow-large" }
    if ($MaxMB -gt 0) { $wordArguments += @("--max-mb", $MaxMB.ToString([System.Globalization.CultureInfo]::InvariantCulture)) }
    if ($SplitDir) { $wordArguments += @("--split-dir", $SplitDir) }
    if ($SplitPrefix) { $wordArguments += @("--split-prefix", $SplitPrefix) }
    if ($AllowLargeAttachments) { $wordArguments += "--allow-large-attachments" }
    if ($MergeWordAttachments) { $wordArguments += "--merge-word-attachments" }
    & $wordPython.Source @wordArguments
    $wordExitCode = $LASTEXITCODE
} finally {
    Pop-Location
}
exit $wordExitCode

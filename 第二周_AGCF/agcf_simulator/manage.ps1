param([ValidateSet('start','stop','status')][string]$Action = 'start', [switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
$appRoot = $PSScriptRoot
$pidFile = Join-Path $appRoot 'server.pid'
$serverPath = Join-Path $appRoot 'server.py'
$url = 'http://127.0.0.1:18766'

function Get-AppProcess {
    try {
        $health = Invoke-RestMethod -Uri "$url/api/health" -TimeoutSec 2
        if ($health.app -eq 'agcf-offline-lab' -and $health.workspace -eq $appRoot) {
            $process = Get-Process -Id $health.pid -ErrorAction SilentlyContinue
            if ($process -and $process.ProcessName -match '^python(w)?$') { return $process }
        }
    } catch { }
    return $null
}

$existingProcess = Get-AppProcess
if ($Action -eq 'stop') {
    if ($existingProcess) { Stop-Process -Id $existingProcess.Id; $existingProcess.WaitForExit(3000) | Out-Null; Write-Host 'AGCF simulator stopped.' }
    else { Write-Host 'AGCF simulator is not running.' }
    if (Test-Path -LiteralPath $pidFile) { Remove-Item -LiteralPath $pidFile }
    exit
}
if ($Action -eq 'status') {
    if ($existingProcess) { Write-Host "Running: $url (PID $($existingProcess.Id))" }
    else { Write-Host 'Not running.' }
    exit
}
if (-not $existingProcess) {
    $bundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $bundledPython) { $pythonExe = $bundledPython }
    else {
        $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
        if (-not $pythonCommand) { throw 'Python 3.10+ is required.' }
        $pythonExe = $pythonCommand.Source
    }
    $portOccupied = $false
    $portProbe = New-Object System.Net.Sockets.TcpClient
    try { $portProbe.Connect('127.0.0.1', 18766); $portOccupied = $true } catch { }
    finally { $portProbe.Dispose() }
    if ($portOccupied) { throw 'Port 18766 is occupied. No process was stopped. Check the existing service first.' }
    $arguments = @('-u', ('"' + $serverPath + '"'), '--port', '18766')
    $newProcess = Start-Process -FilePath $pythonExe -ArgumentList $arguments -WorkingDirectory $appRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $appRoot 'server.log') -RedirectStandardError (Join-Path $appRoot 'server-error.log')
    $newProcess.Id | Set-Content -LiteralPath $pidFile -Encoding ascii
    $ready = $false
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        Start-Sleep -Milliseconds 250
        try {
            $health = Invoke-RestMethod -Uri "$url/api/health" -TimeoutSec 1
            if ($health.app -eq 'agcf-offline-lab' -and $health.pid -eq $newProcess.Id -and $health.workspace -eq $appRoot) { $ready = $true; break }
        } catch { }
    }
    if (-not $ready) { throw 'Server did not become ready. See server-error.log.' }
}
Write-Host "AGCF simulator: $url"
if (-not $NoBrowser) { Start-Process $url }

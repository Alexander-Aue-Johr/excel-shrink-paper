param(
    [string]$DataRoot = 'F:\corpora\enron-v2',
    [string]$Python = 'python',
    [string]$StudyRoot = 'F:\repos\excel-shrink-rust\.tmp\regression\study',
    [switch]$RetryReview
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
New-Item -ItemType Directory -Force -Path $DataRoot | Out-Null
$stamp = [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')
$run = Join-Path $DataRoot "launches\$stamp"
$snapshot = Join-Path $run 'code'
New-Item -ItemType Directory -Force -Path "$snapshot\src", "$snapshot\scripts", "$snapshot\experiments\enron" | Out-Null
# A running download keeps the code it started with, even if the checkout changes.
Copy-Item -LiteralPath "$repo\src\excel_shrink_research" -Destination "$snapshot\src" -Recurse
Copy-Item -LiteralPath "$PSScriptRoot\enron.py" -Destination "$snapshot\scripts\enron.py"
Copy-Item -LiteralPath "$repo\experiments\enron\config.json" -Destination "$snapshot\experiments\enron\config.json"
$files = @{}
Get-ChildItem -LiteralPath $snapshot -File -Recurse | Where-Object { $_.Extension -ne '.pyc' } | ForEach-Object {
    $relative = $_.FullName.Substring($snapshot.Length + 1).Replace('\', '/')
    $files[$relative] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
}
$commit = & git -c "safe.directory=$($repo.Replace('\','/'))" -C $repo rev-parse HEAD
if ($LASTEXITCODE -ne 0) { throw 'Could not record the repository commit.' }
$dirty = & git -c "safe.directory=$($repo.Replace('\','/'))" -C $repo status --porcelain
$record = [ordered]@{ commit = $commit.Trim(); dirty = [bool]$dirty; files = $files }
$recordPath = Join-Path $run 'code.json'
$record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $recordPath -Encoding UTF8
if ((Test-Path -LiteralPath "$StudyRoot\prepared") -and -not (Test-Path -LiteralPath "$DataRoot\reuse-export.stop")) {
    $exportRunning = $false
    $exportReceipt = Join-Path $DataRoot 'reuse-export-process.json'
    if (Test-Path -LiteralPath $exportReceipt) {
        $receipt = Get-Content -LiteralPath $exportReceipt -Raw | ConvertFrom-Json
        $existing = Get-Process -Id $receipt.pid -ErrorAction SilentlyContinue
        if ($existing) {
            $expectedStart = [DateTimeOffset]::Parse($receipt.started_utc).UtcDateTime
            $exportRunning = ($existing.StartTime.ToUniversalTime() -eq $expectedStart)
        }
    }
    if (-not $exportRunning) {
        $exportArgs = @('-u', ('"' + "$snapshot\src\excel_shrink_research\enron\study_export.py" + '"'), ('"' + $StudyRoot + '"'), ('"' + "$DataRoot\reuse" + '"'), '--watch', '--stop-file', ('"' + "$DataRoot\reuse-export.stop" + '"'), '--interval', '15')
        $exporter = Start-Process -FilePath $Python -ArgumentList $exportArgs -WorkingDirectory $snapshot -WindowStyle Hidden -PassThru -RedirectStandardOutput "$DataRoot\reuse-export.log" -RedirectStandardError "$DataRoot\reuse-export.stderr.log"
        [ordered]@{ pid = $exporter.Id; started_utc = $exporter.StartTime.ToUniversalTime().ToString('o') } | ConvertTo-Json | Set-Content -LiteralPath $exportReceipt -Encoding UTF8
        Write-Output "Study preservation started (PID $($exporter.Id))."
    }
}
$arguments = @('-u', ('"' + "$snapshot\scripts\enron.py" + '"'), '--root', ('"' + $DataRoot + '"'), 'prepare', '--code-record', ('"' + $recordPath + '"'))
if ($RetryReview) { $arguments += '--retry-review' }
$process = Start-Process -FilePath $Python -ArgumentList $arguments -WorkingDirectory $snapshot -WindowStyle Hidden -PassThru -RedirectStandardOutput "$run\stdout.log" -RedirectStandardError "$run\stderr.log"
[ordered]@{ pid = $process.Id; started_utc = $process.StartTime.ToUniversalTime().ToString('o'); logs = $run } | ConvertTo-Json | Set-Content -LiteralPath "$DataRoot\launcher.json" -Encoding UTF8
Write-Output "Enron preparation started (PID $($process.Id))."
Write-Output "Logs: $run"
Write-Output "Status: python scripts/enron.py --root `"$DataRoot`" status"

param(
    [Parameter(Mandatory=$true)][string]$StudyRoot,
    [switch]$SmokeTest,
    [string]$Python = (Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe')
)
$ErrorActionPreference = 'Stop'
$StudyRoot = [IO.Path]::GetFullPath($StudyRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)
if (!(Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw 'Python was not found. Supply its executable path with -Python.'
}
$python = (Resolve-Path -LiteralPath $Python).Path
$resultName = if ($SmokeTest) { 'standalone-smoke-pipeline-4096mib' } else { 'standalone-pipeline-4096mib' }
$resultRoot = Join-Path $StudyRoot $resultName
New-Item -ItemType Directory -Path $resultRoot -Force | Out-Null
$software = Join-Path $resultRoot 'software'
if (!(Test-Path -LiteralPath $software)) {
    New-Item -ItemType Junction -Path $software -Target (Join-Path $StudyRoot 'software') | Out-Null
}
$manifest = Join-Path $StudyRoot 'manifest.json'
$binary = Join-Path (Split-Path $StudyRoot -Parent) 'sheetshrink-top49-repeated\bin\sheetshrink.exe'
if (!(Test-Path -LiteralPath $binary)) { throw "SheetShrink binary is missing at $binary" }
$repetitions = 3
if ($SmokeTest) {
    $manifest = Join-Path $StudyRoot 'probe.json'
    $probe = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
    foreach ($case in $probe.cases) {
        $original = $case.variants | Where-Object tool -eq 'Original'
        $case.variants += @{tool='SheetShrink';path=$original.path;sha256=$original.sha256;cleanup_seconds=$null}
    }
    $manifest = Join-Path $resultRoot 'smoke-manifest.json'
    $probe | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -LiteralPath $manifest
    $repetitions = 1
}
$processes = Get-CimInstance Win32_Process | Select-Object Name,ProcessId,ParentProcessId,WorkingSetSize,PrivatePageCount
$available = Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize,FreePhysicalMemory,TotalVirtualMemorySize,FreeVirtualMemory
@{
    started_utc = [DateTime]::UtcNow.ToString('o')
    computer_name = $env:COMPUTERNAME
    processes = $processes
    memory = $available
    sheetshrink_binary_sha256 = (Get-FileHash -LiteralPath $binary -Algorithm SHA256).Hash.ToLower()
    sheetshrink_workers = [Environment]::ProcessorCount
    limit_mib = 4096
    reserve_mib = 0
    timeout_seconds = 900
    repetitions = $repetitions
    manifest_sha256 = (Get-FileHash -LiteralPath $manifest -Algorithm SHA256).Hash.ToLower()
} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 (Join-Path $resultRoot ('environment-' + [DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss') + '.json'))
Start-Transcript -Path (Join-Path $resultRoot 'console.log') -Append | Out-Null
try {
    Write-Host 'Aggregate job commit limit is 4096 MiB. Free physical RAM reserve is disabled.'
    & $python -u (Join-Path $PSScriptRoot 'run_bounded.py') --root $resultRoot --manifest $manifest --helper-directory (Join-Path $StudyRoot 'software\helpers') --repetitions $repetitions --limit-mib 4096 --reserve-mib 0 --timeout 900 --sheetshrink-binary $binary
    if ($LASTEXITCODE -ne 0) { throw "Benchmark exited with code $LASTEXITCODE" }
    if (!$SmokeTest) {
        $csvManifest = Join-Path $resultRoot 'wanderungen-csv.json'
        $selection = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json
        @{cases = @($selection.cases | Where-Object case_id -eq 'wanderungen')} | ConvertTo-Json -Depth 10 | Set-Content -Encoding UTF8 -LiteralPath $csvManifest
        & $python -u (Join-Path $PSScriptRoot 'run_bounded.py') --root $resultRoot --manifest $csvManifest --helper-directory (Join-Path $StudyRoot 'software\helpers') --libraries pandas-csv --repetitions $repetitions --limit-mib 4096 --reserve-mib 0 --timeout 900 --sheetshrink-binary $binary
        if ($LASTEXITCODE -ne 0) { throw "CSV benchmark exited with code $LASTEXITCODE" }
        $env:PYTHONPATH = 'C:\texlive\python-packages'
        & $python (Join-Path $PSScriptRoot 'plot_results.py') $resultRoot
        if ($LASTEXITCODE -ne 0) { throw "Plot generation exited with code $LASTEXITCODE" }
    }
    Write-Host "Finished. Results are in $resultRoot"
}
finally {
    Stop-Transcript | Out-Null
}

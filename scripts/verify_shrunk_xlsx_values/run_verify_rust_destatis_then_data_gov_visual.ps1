param(
    [int]$WorkbookTimeoutSec = 300,
    [int]$WorkbookTimeoutMaxSec = 600,
    [int]$WorkbookTimeoutAttempts = 3,
    [int]$VisualRows = 60,
    [int]$VisualCols = 20,
    [int]$VisualMaxSheets = 0,
    [int]$VisualPixelTolerance = 0,
    [switch]$NoRecompareAll
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..")
Set-Location $repoRoot

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$verifyScript = Join-Path $repoRoot "scripts\verify_shrunk_xlsx_values\verify_shrunk_xlsx_values.py"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path $python)) {
    throw "Python venv not found: $python"
}
if (-not (Test-Path $verifyScript)) {
    throw "Verify script not found: $verifyScript"
}

function Invoke-VerifyRun {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Originals,
        [Parameter(Mandatory = $true)][string]$Shrunk,
        [Parameter(Mandatory = $true)][string]$Out,
        [Parameter(Mandatory = $true)][string]$Log,
        [Parameter(Mandatory = $true)][string]$Screenshots,
        [string]$RepairOriginalsDir = ""
    )

    New-Item -ItemType Directory -Force -Path (Split-Path $Out) | Out-Null
    New-Item -ItemType Directory -Force -Path (Split-Path $Log) | Out-Null
    New-Item -ItemType Directory -Force -Path $Screenshots | Out-Null

    $argsList = @(
        $verifyScript,
        $Originals,
        $Shrunk,
        "--out", $Out,
        "--log-file", $Log,
        "--visual-compare",
        "--visual-rows", "$VisualRows",
        "--visual-cols", "$VisualCols",
        "--visual-max-sheets", "$VisualMaxSheets",
        "--visual-pixel-tolerance", "$VisualPixelTolerance",
        "--visual-screenshot-dir", $Screenshots,
        "--workbook-timeout-sec", "$WorkbookTimeoutSec",
        "--workbook-timeout-max-sec", "$WorkbookTimeoutMaxSec",
        "--workbook-timeout-attempts", "$WorkbookTimeoutAttempts"
    )

    if (-not $NoRecompareAll) {
        $argsList += "--recompare-all"
    }
    if ($RepairOriginalsDir) {
        $argsList += @("--repair-originals-dir", $RepairOriginalsDir)
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting $Name"
    Write-Host "  log:    $Log"
    Write-Host "  report: $Out"
    Write-Host "  shots:  $Screenshots"

    & $python @argsList
    $exitCode = $LASTEXITCODE

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Finished $Name with exit code $exitCode"
    if ($exitCode -ne 0) {
        throw "$Name failed with exit code $exitCode"
    }
}

$destatisOut = "scripts\analyze_destatis_xlsx_files\report_rust_visual_$stamp.json"
$destatisLog = "scripts\analyze_destatis_xlsx_files\verify_destatis_rust_visual_$stamp.log"
$destatisScreenshots = "scripts\analyze_destatis_xlsx_files\visual_screenshots_$stamp"

$dataGovOut = "scripts\analyze_data_gov_xlsx_files\report_rust_visual_$stamp.json"
$dataGovLog = "scripts\analyze_data_gov_xlsx_files\verify_data_gov_rust_visual_$stamp.log"
$dataGovScreenshots = "scripts\analyze_data_gov_xlsx_files\visual_screenshots_$stamp"

Invoke-VerifyRun `
    -Name "Destatis rust verify with visual compare" `
    -Originals "scripts\scrape_destatis_xlsx_files\downloaded" `
    -Shrunk "scripts\analyze_destatis_xlsx_files\shrinked_by_excel_shrink_rust" `
    -Out $destatisOut `
    -Log $destatisLog `
    -Screenshots $destatisScreenshots `
    -RepairOriginalsDir "scripts\analyze_destatis_xlsx_files\fixed_original_corrupt_destatis_files"

Invoke-VerifyRun `
    -Name "data.gov rust verify with visual compare" `
    -Originals "scripts\scrape_data_gov_xlsx_files\downloaded" `
    -Shrunk "scripts\analyze_data_gov_xlsx_files\shrinked_by_excel_shrink_rust" `
    -Out $dataGovOut `
    -Log $dataGovLog `
    -Screenshots $dataGovScreenshots

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] All verify runs finished."

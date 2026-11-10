param(
    [Parameter(Mandatory=$true)][string]$StudyRoot,
    [Parameter(Mandatory=$true)][string]$Binary,
    [Parameter(Mandatory=$true)][string]$EngineSource,
    [Parameter(Mandatory=$true)][string]$Output,
    [ValidateSet('nxpowerlite','check-performance','xlstyles','sheetshrink')]
    [string[]]$Tools = @('nxpowerlite','check-performance','xlstyles','sheetshrink'),
    [string[]]$Cases = @('all'),
    [ValidateRange(1,86400)][int]$OperationTimeoutSeconds = 300,
    [ValidateRange(1,86400)][int]$IdleTimeoutSeconds = 300
)
$ErrorActionPreference = 'Stop'

function Save-Json($Path, $Value) {
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}

function Get-SelectedWorkbooks($Panel) {
    if ($Cases -contains 'all') { return @($Panel.workbooks) }
    $selected = @($Panel.workbooks | Where-Object { $Cases -contains $_.id })
    if ($selected.Count -ne $Cases.Count) { throw 'Every selected case must match a frozen panel ID.' }
    return $selected
}

function Assert-Inputs($Workbooks) {
    foreach ($workbook in $Workbooks) {
        $original = Join-Path $StudyRoot "inputs\$($workbook.file)"
        if ((Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash.ToLowerInvariant() -ne $workbook.input_sha256) {
            throw "Original identity mismatch for $($workbook.file)"
        }
        foreach ($tool in $Tools) {
            $candidate = Join-Path $StudyRoot "comparators\$tool\outputs\$($workbook.file)"
            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "Missing output $candidate" }
        }
    }
}

function Compare-ToolOutputs($Tool, $Workbooks, $Root) {
    $frame = Join-Path $Root "frames\$Tool"
    $originals = Join-Path $frame 'originals'
    $candidates = Join-Path $frame 'candidates'
    [void][IO.Directory]::CreateDirectory($originals)
    [void][IO.Directory]::CreateDirectory($candidates)
    $identities = @()
    foreach ($workbook in $Workbooks) {
        Copy-Item -LiteralPath (Join-Path $StudyRoot "inputs\$($workbook.file)") -Destination $originals
        $candidate = Join-Path $StudyRoot "comparators\$Tool\outputs\$($workbook.file)"
        Copy-Item -LiteralPath $candidate -Destination $candidates
        $identities += @{file=$workbook.file; input_sha256=$workbook.input_sha256;
            candidate_sha256=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash.ToLowerInvariant()}
    }
    $config = (& $Binary template | Out-String | ConvertFrom-Json)
    $config.source_a = $originals
    $config.source_b = $candidates
    $config.report_dir = Join-Path $Root "reports\$Tool"
    $config.max_processes = 1
    $config.continue_on_error = $true
    $config.capture.operation_timeout_seconds = $OperationTimeoutSeconds
    $config.capture.idle_timeout_seconds = $IdleTimeoutSeconds
    $config.capture.retain_equal_images = $true
    $configuration = Join-Path $Root "$Tool-config.json"
    Save-Json $configuration $config
    $arguments = @('run', '--config', ('"{0}"' -f $configuration))
    $process = Start-Process -FilePath $Binary -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru `
        -RedirectStandardOutput (Join-Path $Root "$Tool-stdout.log") `
        -RedirectStandardError (Join-Path $Root "$Tool-stderr.log")
    return @{tool=$Tool; exit_code=$process.ExitCode; inputs=$identities; configuration=$configuration;
        report=$config.report_dir; finished_at=[DateTime]::UtcNow.ToString('o')}
}

$panelPath = Join-Path $PSScriptRoot '../../papers/research-audit/comparator-benchmark-20261109/panel.json'
$panel = Get-Content -LiteralPath $panelPath -Raw | ConvertFrom-Json
$workbooks = Get-SelectedWorkbooks $panel
$StudyRoot = (Resolve-Path -LiteralPath $StudyRoot).Path
$Binary = (Resolve-Path -LiteralPath $Binary).Path
$EngineSource = (Resolve-Path -LiteralPath $EngineSource).Path
$root = [IO.Path]::GetFullPath($Output)
if (Test-Path -LiteralPath $root) { throw 'Choose a new validation output directory.' }
Assert-Inputs $workbooks
[void][IO.Directory]::CreateDirectory($root)
$env:EXCEL_SHRINK_CAPTURE_DIR = Join-Path $EngineSource 'excel-shrink-cli\office-capture'
$env:EXCEL_SHRINK_LOG_DIR = Join-Path $root 'logs'
$env:TEMP = Join-Path $root 'temp'
$env:TMP = $env:TEMP
$env:RAYON_NUM_THREADS = '1'
[void][IO.Directory]::CreateDirectory($env:TEMP)
$receipt = @{operation='ordinary DeltaSieve comparison'; source_commit=$panel.source_commit;
    binary_sha256=(Get-FileHash -LiteralPath $Binary -Algorithm SHA256).Hash.ToLowerInvariant();
    rayon_threads=1; trials=@()}
foreach ($tool in $Tools) {
    $receipt.trials += Compare-ToolOutputs $tool $workbooks $root
    Save-Json (Join-Path $root 'execution.json') $receipt
}
if (@($receipt.trials | Where-Object { $_.exit_code -ne 0 }).Count -gt 0) {
    throw "One or more comparisons failed. Inspect the retained reports and $root\execution.json."
}

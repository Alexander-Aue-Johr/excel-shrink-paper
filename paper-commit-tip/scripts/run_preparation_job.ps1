[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$JobFile)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$job = [IO.File]::ReadAllText([IO.Path]::GetFullPath($JobFile)) | ConvertFrom-Json

function Get-RequiredText([string]$Name) {
    $property = $job.PSObject.Properties[$Name]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        throw "The preparation job is missing $Name."
    }
    return [string]$property.Value
}

$repo = [IO.Path]::GetFullPath((Get-RequiredText 'rust_repo'))
$baseConfiguration = [IO.Path]::GetFullPath((Get-RequiredText 'base_configuration'))
$configurationPath = [IO.Path]::GetFullPath((Get-RequiredText 'configuration_path'))
$runDirectory = [IO.Path]::GetFullPath((Get-RequiredText 'run_directory'))
$executable = [IO.Path]::GetFullPath((Get-RequiredText 'executable'))
$receiptPath = [IO.Path]::GetFullPath((Get-RequiredText 'receipt_path'))
$outputRoot = [IO.Path]::GetFullPath((Get-RequiredText 'output_root'))
$corpus = Get-RequiredText 'corpus'
$expectedHash = Get-RequiredText 'executable_sha256'
if ($expectedHash -notmatch '\A[0-9a-fA-F]{64}\z' -or
    (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash -ine $expectedHash) {
    throw 'The preparation executable differs from the reviewed build. No run was started.'
}
if ($configurationPath.Equals($baseConfiguration, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'A preparation job must save its own configuration, separate from the base configuration.'
}
if (-not [IO.File]::Exists($baseConfiguration)) { throw 'The base launcher configuration is missing.' }

Import-Module (Join-Path $repo 'tools/corpus-launcher/Core.psm1') -Force
Import-Module (Join-Path $repo 'tools/corpus-launcher/Launch.psm1') -Force
$configuration = Import-StudyConfiguration $repo $baseConfiguration
$configuration.settings.corpus = $corpus
$configuration.settings.output_root = $outputRoot
$configuration.settings.control_room_port = Get-RequiredText 'port'
$configuration.settings.partition = 0
$configuration.settings.comparison_groups = 1
$configuration.settings.hide_excel = $true
$configuration.settings.build_release = $false
$configuration.settings.order = 'size-asc'
$configuration.settings.import_report = $null
$configuration.settings.reuse_prepared_from = @((Get-Value $job 'approved_builds' @()) | Select-Object -Unique)

$archive = Get-Value $job 'archive'
if ($archive) {
    $corpusId = Get-RequiredText 'corpus_id'
    if ($corpusId -notmatch '\A[0-9a-f]{64}\z') { throw 'The explicit corpus ID must be a lowercase SHA-256 value.' }
    $total = 0L
    if (-not [long]::TryParse((Get-RequiredText 'total_workbooks'), [ref]$total) -or $total -lt 1) {
        throw 'The explicit corpus workbook count must be positive.'
    }
    $configuration.corpora = @([pscustomobject]@{
        name=$corpus; id=$corpusId; local_archive=[IO.Path]::GetFullPath([string]$archive)
        url=$null; total_workbooks=$total
    })
}
$retrySkipped = Get-Value $job 'retry_skipped' $false
if ($retrySkipped -isnot [bool]) { throw 'retry_skipped must be a JSON boolean.' }
$plan = New-StudyLaunchPlan $configuration $repo $executable $runDirectory -PrepareOnly
if ($retrySkipped) {
    $plan.Arguments += '--retry-skipped'
    $plan.CommandLine = ($plan.Arguments | ForEach-Object { ConvertTo-WindowsArgument ([string]$_) }) -join ' '
}
$receipt = Start-StudyLaunch $configuration $configurationPath $repo $plan
Save-StudyJson $receiptPath $receipt
Write-Output "Preparation started for $corpus (PID $($receipt.Pid))."
Write-Output "Log: $($receipt.Stderr)"

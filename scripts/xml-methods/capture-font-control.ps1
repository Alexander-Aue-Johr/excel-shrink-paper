<#
.SYNOPSIS
Capture the three reduced font-witness controls with installed Excel.
#>
param(
    [string]$Directory = (
        Join-Path $PSScriptRoot '../../papers/sheetshrink/evidence/xml-methods/native-font'
    )
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $Directory).Path
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
$results = @()

try {
    foreach ($name in @('source', 'one-witness', 'no-witness')) {
        $inputPath = Join-Path $root ($name + '.xlsx')
        $book = $excel.Workbooks.Open($inputPath, 0, $true)
        try {
            $excel.CalculateFullRebuild()
            $sheet = $book.Worksheets.Item(1)
            $height = $sheet.Rows.Item(2).RowHeight
            $range = $sheet.Range('A1:D4')
            # CopyPicture's screen bitmap is pasted into a disposable chart for export.
            $range.CopyPicture(1, 2)
            $chart = $sheet.ChartObjects().Add(0, 100, $range.Width, $range.Height)
            [void]$chart.Activate()
            [void]$chart.Chart.Paste()
            $imageName = $name + '.png'
            $exported = $chart.Chart.Export((Join-Path $root $imageName), 'PNG')
            if (-not $exported) {
                throw ('Excel could not export ' + $imageName)
            }
            $results += @{
                variant = $name
                row_2_height_points = $height
                image = $imageName
                exported = $exported
                excel_version = $excel.Version
                excel_build = $excel.Build
            }
            Write-Output ('Captured ' + $name + '. Row 2 height is ' + $height + ' pt.')
        } finally {
            [void]$book.Close($false)
        }
    }
    $results | ConvertTo-Json -Depth 5 |
        Set-Content -Encoding utf8 (Join-Path $root 'receipt.json')
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

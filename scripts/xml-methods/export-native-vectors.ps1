<#
.SYNOPSIS
Export the retained controls with Excel's native print renderer as vector PDFs.
#>
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '../../tmp/pdfs/xml-methods/vector-export'),
    [string]$CorpusCaseDirectory = 'F:\corpora\paper-cases\enron-fill-override-ccde4294'
)
$ErrorActionPreference = 'Stop'
$fontRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../papers/sheetshrink/evidence/xml-methods/native-font')).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$corpusRoot = Join-Path $CorpusCaseDirectory 'rebuilt'
$plans = @(
    @{group='font'; name='source'; input=(Join-Path $fontRoot 'source.xlsx'); sheet='Font witness'; range='A1:D4'},
    @{group='font'; name='one-witness'; input=(Join-Path $fontRoot 'one-witness.xlsx'); sheet='Font witness'; range='A1:D4'},
    @{group='font'; name='no-witness'; input=(Join-Path $fontRoot 'no-witness.xlsx'); sheet='Font witness'; range='A1:D4'},
    @{group='enron'; name='Original'; input=(Join-Path $corpusRoot 'Original.xlsx'); sheet='PL-Consolidated'; range='A25:H37'},
    @{group='enron'; name='Default'; input=(Join-Path $corpusRoot 'Default.xlsx'); sheet='PL-Consolidated'; range='A25:H37'},
    @{group='enron'; name='Check-Performance'; input=(Join-Path $corpusRoot 'Check-Performance.xlsx'); sheet='PL-Consolidated'; range='A25:H37'}
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3
$receipts = @()
try {
    foreach ($plan in $plans) {
        $book = $excel.Workbooks.Open($plan.input, 0, $true)
        try {
            $excel.CalculateFullRebuild()
            $sheet = $book.Worksheets.Item($plan.sheet)
            $range = $sheet.Range($plan.range)
            $setup = $sheet.PageSetup
            $setup.PrintArea = $range.Address()
            $setup.PrintTitleRows = ''
            $setup.PrintTitleColumns = ''
            $setup.LeftMargin = 0
            $setup.RightMargin = 0
            $setup.TopMargin = 0
            $setup.BottomMargin = 0
            $setup.HeaderMargin = 0
            $setup.FooterMargin = 0
            $setup.LeftHeader = ''
            $setup.CenterHeader = ''
            $setup.RightHeader = ''
            $setup.LeftFooter = ''
            $setup.CenterFooter = ''
            $setup.RightFooter = ''
            $setup.CenterHorizontally = $false
            $setup.CenterVertically = $false
            $setup.PrintGridlines = $true
            $setup.PrintHeadings = $false
            $setup.Orientation = 2
            $setup.PaperSize = 8
            $setup.Zoom = 100
            $pdfName = $plan.group + '-' + $plan.name + '.pdf'
            $pdf = Join-Path $OutputDirectory $pdfName
            $sheet.ExportAsFixedFormat(0, $pdf, 0, $false, $false)
            $receipts += @{
                group = $plan.group
                variant = $plan.name
                input_sha256 = (Get-FileHash -LiteralPath $plan.input -Algorithm SHA256).Hash.ToLower()
                sheet = $plan.sheet
                range = $plan.range
                pdf = $pdfName
                width_points = $range.Width
                height_points = $range.Height
                row_2_height_points = $sheet.Rows.Item(2).RowHeight
                excel_version = $excel.Version
                excel_build = $excel.Build
                method = 'Read-only open, CalculateFullRebuild, fixed-format PDF export at 100 percent, zero margins, printed gridlines, no Save'
            }
            Write-Output ('Exported ' + $pdfName)
        } finally {
            [void]$book.Close($false)
        }
    }
    $receipts | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $OutputDirectory 'receipt.json')
} finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

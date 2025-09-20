<#!
.SYNOPSIS
    Convert an Excel workbook (e.g., BattleRhythm.xlsx) to CSV.

.DESCRIPTION
    Uses Excel COM automation (built into Windows with Microsoft Excel installed) to export
    either a single worksheet to a CSV file or all worksheets to individual CSV files.

    - By default, exports the first visible worksheet to <InputBaseName>.csv
    - Use -SheetName or -SheetIndex to choose a specific sheet
    - Use -AllSheets to export every (visible) sheet to a folder

.PARAMETER InputPath
    Path to the Excel workbook (.xlsx, .xlsm, .xls). Defaults to ./BattleRhythm.xlsx

.PARAMETER OutputPath
    For single-sheet export: output CSV file path
    For -AllSheets: output directory (it will be created if missing)

.PARAMETER SheetName
    Name of the worksheet to export (single-sheet mode)

.PARAMETER SheetIndex
    1-based index of the worksheet to export (single-sheet mode)

.PARAMETER AllSheets
    Export all sheets, each to its own CSV in an output folder

.PARAMETER IncludeHidden
    Include hidden/very hidden worksheets when using -AllSheets

.PARAMETER Overwrite
    Overwrite existing CSV files if present

.EXAMPLE
    # Export first visible sheet to BattleRhythm.csv (in same folder)
    .\convert-BattleRhythm.ps1 -InputPath .\BattleRhythm.xlsx

.EXAMPLE
    # Export a specific sheet to a specific CSV file
    .\convert-BattleRhythm.ps1 -InputPath .\BattleRhythm.xlsx -SheetName "Schedule" -OutputPath .\schedule.csv

.EXAMPLE
    # Export all visible sheets to a new folder named BattleRhythm-csv
    .\convert-BattleRhythm.ps1 -InputPath .\BattleRhythm.xlsx -AllSheets

.NOTES
    Requires Microsoft Excel to be installed on this machine.
#>

[CmdletBinding()] Param(
    [Parameter(Position=0)]
    [string]$InputPath = "./BattleRhythm.xlsx",

    [Parameter(Position=1)]
    [string]$OutputPath,

    [string]$SheetName,
    [int]$SheetIndex,
    [switch]$AllSheets,
    [switch]$IncludeHidden,
    [switch]$Overwrite
)

function Resolve-FullPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try { return (Resolve-Path -Path $Path -ErrorAction Stop).Path } catch { return (Join-Path -Path (Get-Location) -ChildPath $Path) }
}

function Get-CsvFormatCode {
    # Prefer UTF-8 CSV (62) if supported; otherwise fall back to legacy CSV (6)
    $xlCSVUTF8 = 62
    $xlCSV = 6
    return @{ Preferred=$xlCSVUTF8; Fallback=$xlCSV }
}

function Test-Directory([string]$Path) {
    return (Test-Path -LiteralPath $Path -PathType Container)
}

function Ensure-Directory([string]$Path) {
    if (-not (Test-Directory $Path)) { [void](New-Item -ItemType Directory -Path $Path -Force) }
}

function Sanitize-Name([string]$Name) {
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $Name = $Name -replace [Regex]::Escape([string]$c), '_' }
    return $Name
}

function Export-ExcelSheetToCsv {
    param(
        [Parameter(Mandatory)] [__ComObject]$Excel,
        [Parameter(Mandatory)] [__ComObject]$Worksheet,
        [Parameter(Mandatory)] [string]$CsvPath,
        [Parameter(Mandatory)] [int]$FormatCodePreferred,
        [Parameter(Mandatory)] [int]$FormatCodeFallback,
        [switch]$Overwrite
    )

    if ((Test-Path -LiteralPath $CsvPath) -and -not $Overwrite) {
        throw "Output exists: $CsvPath (use -Overwrite to replace)"
    }

    # Copy worksheet to a new temporary workbook to avoid touching the original
    $Worksheet.Copy() | Out-Null
    $tempWb = $Excel.ActiveWorkbook
    try {
        try {
            $tempWb.SaveAs($CsvPath, $FormatCodePreferred)
        } catch {
            # Fallback for older Excel versions
            $tempWb.SaveAs($CsvPath, $FormatCodeFallback)
        }
    } finally {
        $tempWb.Close($false)
        # Release COM object
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($tempWb)
    }
}

function Convert-ExcelToCsv {
    param(
        [Parameter(Mandatory)] [string]$InputFullPath,
        [string]$OutputPath,
        [string]$SheetName,
        [int]$SheetIndex,
        [switch]$AllSheets,
        [switch]$IncludeHidden,
        [switch]$Overwrite
    )

    if (-not (Test-Path -LiteralPath $InputFullPath -PathType Leaf)) {
        throw "Input file not found: $InputFullPath"
    }

    $formats = Get-CsvFormatCode
    $formatPreferred = $formats.Preferred
    $formatFallback = $formats.Fallback

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($InputFullPath)

        if ($AllSheets) {
            # Determine target folder
            $targetDir = $OutputPath
            if ([string]::IsNullOrWhiteSpace($targetDir)) {
                $baseDir = Split-Path -Path $InputFullPath -Parent
                $baseName = [IO.Path]::GetFileNameWithoutExtension($InputFullPath)
                $targetDir = Join-Path $baseDir ("{0}-csv" -f $baseName)
            } elseif (-not (Test-Directory $targetDir)) {
                # If OutputPath is a file-like path (ends with .csv), use its directory
                if ($OutputPath.ToLower().EndsWith('.csv')) {
                    $targetDir = Split-Path -Path $OutputPath -Parent
                }
            }
            Ensure-Directory $targetDir

            foreach ($ws in @($workbook.Worksheets)) {
                try {
                    # Visibility: -1 visible, 0 hidden, 2 very hidden
                    if (-not $IncludeHidden -and $ws.Visible -ne -1) { continue }
                    $sheetName = Sanitize-Name($ws.Name)
                    $csvPath = Join-Path $targetDir ("{0}__{1}.csv" -f ([IO.Path]::GetFileNameWithoutExtension($InputFullPath)), $sheetName)
                    Write-Verbose "Exporting sheet '$sheetName' -> $csvPath"
                    Export-ExcelSheetToCsv -Excel $excel -Worksheet $ws -CsvPath $csvPath -FormatCodePreferred $formatPreferred -FormatCodeFallback $formatFallback -Overwrite:$Overwrite
                } finally {
                    if ($ws) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) }
                }
            }
        }
        else {
            # Determine the sheet
            $ws = $null
            if ($SheetName) {
                $ws = $workbook.Worksheets.Item($SheetName)
            } elseif ($SheetIndex) {
                $ws = $workbook.Worksheets.Item([int]$SheetIndex)
            } else {
                # First visible sheet
                foreach ($s in @($workbook.Worksheets)) { if ($s.Visible -eq -1) { $ws = $s; break } }
                if (-not $ws) { $ws = $workbook.Worksheets.Item(1) }
            }

            # Determine output path
            if ([string]::IsNullOrWhiteSpace($OutputPath)) {
                $dir = Split-Path -Path $InputFullPath -Parent
                $base = [IO.Path]::GetFileNameWithoutExtension($InputFullPath)
                $sheetSuffix = ''
                if ($SheetName) { $sheetSuffix = "__" + (Sanitize-Name $SheetName) }
                elseif ($SheetIndex -gt 0) { $sheetSuffix = "__Sheet$SheetIndex" }
                $OutputPath = Join-Path $dir ("{0}{1}.csv" -f $base, $sheetSuffix)
            } else {
                # If OutputPath is a directory, generate file name inside it
                if (Test-Directory $OutputPath) {
                    $base = [IO.Path]::GetFileNameWithoutExtension($InputFullPath)
                    $sheetSuffix = ''
                    if ($SheetName) { $sheetSuffix = "__" + (Sanitize-Name $SheetName) }
                    elseif ($SheetIndex -gt 0) { $sheetSuffix = "__Sheet$SheetIndex" }
                    $OutputPath = Join-Path $OutputPath ("{0}{1}.csv" -f $base, $sheetSuffix)
                }
            }

            $sheetNameLog = $ws.Name
            Write-Verbose "Exporting sheet '$sheetNameLog' -> $OutputPath"
            Export-ExcelSheetToCsv -Excel $excel -Worksheet $ws -CsvPath $OutputPath -FormatCodePreferred $formatPreferred -FormatCodeFallback $formatFallback -Overwrite:$Overwrite
            if ($ws) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ws) }
        }
    } finally {
        if ($workbook) { $workbook.Close($false) | Out-Null; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        if ($excel) { $excel.Quit() | Out-Null; [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect();
    }
}

# Entry point
$inputFull = Resolve-FullPath $InputPath
Convert-ExcelToCsv -InputFullPath $inputFull -OutputPath $OutputPath -SheetName $SheetName -SheetIndex $SheetIndex -AllSheets:$AllSheets -IncludeHidden:$IncludeHidden -Overwrite:$Overwrite

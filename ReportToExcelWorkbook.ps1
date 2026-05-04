# minimally install the ImportExcel module if not present:
# Install-Module ImportExcel -Scope CurrentUser
# You may need to run PowerShell as admin to install modules, but once installed you can run this script without admin privileges.


# -----------------------------------------
# CONFIGURATION
# -----------------------------------------
$ProgramName = "ReportToExcelWorkbook"
$RunDir      = "D:\data\"
$InputFile   = "input.txt"

$RunStamp    = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile     = "$RunDir$ProgramName-$InputFile-$RunStamp.log"
$OutFile     = "$RunDir$ProgramName-$InputFile-$RunStamp.xlsx"

# -----------------------------------------
# LOGGING FUNCTION
# -----------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$timestamp [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

Write-Log "Starting $ProgramName..."
Write-Log "Log file created: $LogFile"

# -----------------------------------------
# READ INPUT FILE
# -----------------------------------------
$FullPath = Join-Path $RunDir $InputFile
$text = Get-Content $FullPath -Raw

# Remove decorative separators
$cleaned = $text -replace "=+\r?\n", ""

# -----------------------------------------
# PARSE INTO SHEETS
# -----------------------------------------
# Regex: capture sheet name + block until next SHEET:
$pattern = "SHEET:\s*(.+?)\r?\n(.*?)(?=\r?\nSHEET:|$)"
$matches = [regex]::Matches($cleaned, $pattern, "Singleline")

$sheets = [ordered] @{}

foreach ($m in $matches) {
    $sheetName = $m.Groups[1].Value.Trim()
    Write-Log "Parsing sheet: $sheetName"

    $block = $m.Groups[2].Value
    $rows = @()

    foreach ($line in $block.Split("`n")) {
        if ($line -notmatch "\|") { continue }

        # Split on pipe, trim whitespace
        $parts = $line.Split("|") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

        # Split on pipe, trim whitespace, remove empties
        # $parts = $line.Split("|") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }


        if ($parts.Count -gt 0) {
            $rows += ,$parts
        }
    }

    if ($rows.Count -gt 0) {
        $sheets[$sheetName] = $rows
    }
    else {
        Write-Log "Sheet '$sheetName' contained no table rows." "WARN"
    }
}

Write-Log "Parsed $($sheets.Keys.Count) sheets successfully."

# -----------------------------------------
# WRITE EXCEL WORKBOOK
# -----------------------------------------
if (Test-Path $OutFile) { Remove-Item $OutFile }

foreach ($sheetName in $sheets.Keys) {

    $rows = $sheets[$sheetName]
    $rowCount = $rows.Count

    # Excel sheet names max 31 chars
    $safeName = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))

    Write-Log "Creating sheet: $safeName with $rowCount rows"

    # Convert rows to objects for Export-Excel
    $header = $rows[0]
    $dataRows = $rows[1..($rows.Count - 1)]

    $objects = foreach ($r in $dataRows) {
        $obj = [ordered]@{}
        for ($i = 0; $i -lt $header.Count; $i++) {
            $colName = $header[$i]
            $value   = if ($i -lt $r.Count) { $r[$i] } else { "" }
            $obj[$colName] = $value
        }
        [pscustomobject]$obj
    }

    # Append sheet to workbook
    $objects | Export-Excel -Path $OutFile -WorksheetName $safeName -AutoSize -Append
}

Write-Log "Workbook saved: $OutFile"
Write-Log "$ProgramName completed successfully."

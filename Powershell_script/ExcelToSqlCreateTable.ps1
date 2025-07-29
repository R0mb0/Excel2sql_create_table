# PowerShell script to generate a SQL CREATE TABLE statement from an Excel file
Write-Host "=== GENERATE CREATE TABLE FROM EXCEL ==="

# --- XLSX File Selection ---
$scriptDir = $PSScriptRoot
$excelFiles = Get-ChildItem -Path $scriptDir -Filter *.xlsx

if (-not $excelFiles -or $excelFiles.Count -eq 0) {
    Write-Error "No .xlsx file found in: $scriptDir"
    exit 1
}

if ($excelFiles.Count -eq 1) {
    $ExcelFile = $excelFiles[0]
    Write-Host "Excel file selected: $($ExcelFile.Name)"
} else {
    Write-Host "Multiple .xlsx files found in the folder:"
    for ($i = 0; $i -lt $excelFiles.Count; $i++) {
        Write-Host "[$i] $($excelFiles[$i].Name)"
    }
    $fileIndex = Read-Host "Enter the number of the Excel file to use"
    if ($fileIndex -notmatch '^\d+$' -or [int]$fileIndex -lt 0 -or [int]$fileIndex -ge $excelFiles.Count) {
        Write-Error "Invalid selection. Exiting."
        exit 1
    }
    $ExcelFile = $excelFiles[$fileIndex]
    Write-Host "Excel file selected: $($ExcelFile.Name)"
}

# --- Sheet Name ---
$SheetName = Read-Host "Enter the Excel sheet name (SheetName)"
if ([string]::IsNullOrWhiteSpace($SheetName)) {
    Write-Error "SheetName not provided. Exiting script."
    exit 1
}

# --- SQL Table Name ---
$SqlTableName = Read-Host "Enter the SQL table name (SqlTableName)"
if ([string]::IsNullOrWhiteSpace($SqlTableName)) {
    Write-Error "SqlTableName not provided. Exiting script."
    exit 1
}

# --- Output File Name ---
$defaultOutput = "create_table.txt"
$outputInput = Read-Host "Enter the output file name (press Enter for default: $defaultOutput)"
if ([string]::IsNullOrWhiteSpace($outputInput)) {
    $OutputFile = Join-Path $PSScriptRoot $defaultOutput
} else {
    $OutputFile = Join-Path $PSScriptRoot $outputInput
}

# --- Type Detection Threshold ---
$defaultThreshold = 500
$thresholdInput = Read-Host "Enter threshold for type detection (press Enter for default: $defaultThreshold)"
if ([string]::IsNullOrWhiteSpace($thresholdInput)) {
    $TypeThreshold = $defaultThreshold
} else {
    if ($thresholdInput -as [int]) {
        $TypeThreshold = [int]$thresholdInput
    } else {
        Write-Error "Invalid threshold. Using default value: $defaultThreshold"
        $TypeThreshold = $defaultThreshold
    }
}

Write-Host "Excel file: $($ExcelFile.FullName)"
Write-Host "Sheet: $SheetName"
Write-Host "SQL table: $SqlTableName"
Write-Host "Output: $OutputFile"

# --- ImportExcel Module ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# --- Read Data from Sheet ---
try {
    $data = Import-Excel -Path $ExcelFile.FullName -WorksheetName $SheetName
} catch {
    Write-Error "Error importing XLSX file: $($_.Exception.Message)"
    exit 1
}
if (-not $data -or $data.Count -eq 0) {
    Write-Error "Sheet is empty or unreadable."
    exit 1
}

# --- Get Actual Column Names ---
$columnNames = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

# --- Handle Duplicate Headers ---
$finalHeaders = @()
$uniqueHeaders = @{}
$duplicates = @()
foreach ($c in $columnNames) {
    $colName = if ($null -eq $c -or "$c".Trim() -eq "") { "UnnamedColumn" } else { "$c".Trim() }
    $baseName = $colName
    $i = 2
    while ($finalHeaders -contains $colName) { 
        if (-not ($duplicates -contains $baseName)) { $duplicates += $baseName }
        $colName = "${baseName}_$i"; $i++ 
    }
    $finalHeaders += $colName
    $uniqueHeaders[$finalHeaders.Count - 1] = $colName
}
if ($duplicates.Count -gt 0) {
    Write-Host "WARNING: Duplicate columns found and renamed: $($duplicates -join ', ')"
} else {
    Write-Host "No duplicate columns detected."
}

function Get-SqlTypeAndLength($colName, $values) {
    # Clean: convert all to string, remove nulls and empty/whitespace strings
    $cleanValues = $values | Where-Object { $_ -ne $null } | ForEach-Object { "$_".Trim() }
    $nonNullValues = $cleanValues | Where-Object { $_ -match '\S' }

    if ($nonNullValues.Count -eq 0) {
        if ($colName -match '(?i)data|date') { return @{Type="DATETIME"; Length=$null} }
        else { return @{Type="NVARCHAR"; Length=100} }
    }

    $intCount = ($nonNullValues | Where-Object { $_ -match '^-?\d+$' } | Measure-Object).Count
    $floatCount = ($nonNullValues | Where-Object { $_ -match '^-?\d+\.\d+$' } | Measure-Object).Count
    $dateCount = ($nonNullValues | Where-Object { 
        try { [datetime]::Parse($_) | Out-Null; $true } catch { $false }
    } | Measure-Object).Count
    $boolCount = ($nonNullValues | Where-Object { 
        $t = $_.ToLower()
        $t -eq "true" -or $t -eq "false" -or $t -eq "0" -or $t -eq "1"
    } | Measure-Object).Count

    if ($intCount -ge $TypeThreshold -and $intCount -eq $nonNullValues.Count) {
        return @{Type="INT"; Length=$null}
    } elseif ($boolCount -ge $TypeThreshold -and $boolCount -eq $nonNullValues.Count) {
        return @{Type="BIT"; Length=$null}
    } elseif ($floatCount -ge $TypeThreshold -and $floatCount -eq $nonNullValues.Count) {
        return @{Type="FLOAT"; Length=$null}
    } elseif ($dateCount -ge $TypeThreshold -and $dateCount -eq $nonNullValues.Count) {
        return @{Type="DATETIME"; Length=$null}
    } else {
        # NVARCHAR: calculate max length
        $maxLen = ($nonNullValues | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
        $maxAllowed = 255
        $colLength = [Math]::Min($maxLen, $maxAllowed)
        return @{Type="NVARCHAR"; Length=$colLength}
    }
}

# --- Analyze Columns and Detect SQL Types ---
$Columns = @()
for ($i = 0; $i -lt $finalHeaders.Count; $i++) {
    $colName = $finalHeaders[$i]
    $propName = $columnNames[$i]
    $values = @()
    foreach ($row in $data) {
        $cellValue = $row.$propName
        $values += $cellValue
    }
    $typeInfo = Get-SqlTypeAndLength $colName $values
    $CleanCol = $colName -replace '[^a-zA-Z0-9_]', '_'
    $sqlType = if ($typeInfo.Type -eq "NVARCHAR") { "NVARCHAR($($typeInfo.Length))" } else { $typeInfo.Type }
    Write-Host "$colName -> $sqlType"
    $Columns += "[$CleanCol] $sqlType"
}

$ColumnsSql = $Columns -join ",`n    "
$CreateTable = "CREATE TABLE [$SqlTableName] (`n    $ColumnsSql`n);"

Set-Content -Path $OutputFile -Value $CreateTable -Encoding UTF8
Write-Host "SQL command generated and saved to: $OutputFile"

# --- Output Preview ---
Write-Host "=== OUTPUT FILE CONTENT: $OutputFile ==="
Get-Content -Path $OutputFile | Write-Host

Write-Host "=== END OF SCRIPT ==="
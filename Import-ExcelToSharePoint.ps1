# Required modules
$requiredModules = @(
    "PnP.PowerShell",
    "ImportExcel"
)

# Function to ensure required modules are installed
function Ensure-ModulesInstalled {
    param (
        [string[]]$ModuleNames
    )
    
    foreach ($module in $ModuleNames) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module..."
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -Force
    }
}

# Function to write formatted header
function Write-Header {
    param (
        [string]$Title
    )
    
    Write-Host "`n═══════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
}

# Function to write formatted section
function Write-Section {
    param (
        [string]$Title
    )
    
    $width = 37
    $line = "─" * ($width - 2)
    Write-Host ""
    Write-Host "┌$line┐" -ForegroundColor DarkCyan
    Write-Host "│ $Title$(" " * ($width - $Title.Length - 3))│" -ForegroundColor Cyan
    Write-Host "└$line┘" -ForegroundColor DarkCyan
    Write-Host ""
}

# Function to write formatted status
function Write-Status {
    param (
        [string]$Message,
        [string]$Status,
        [string]$Color = "White"
    )
    
    $statusIcon = switch ($Status.ToLower()) {
        "success" { "✓"; break }
        "error" { "✗"; break }
        "warning" { "!"; break }
        "info" { "→"; break }
        default { " "; break }
    }
    
    Write-Host "  $statusIcon $Message" -ForegroundColor $Color
}

# Function to write progress to MD file
function Write-ProgressToMD {
    param (
        [string]$Message,
        [string]$Status = "info",
        [bool]$IsRemoved = $false
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $statusIcon = switch ($Status.ToLower()) {
        "success" { "✓"; break }
        "error" { "✗"; break }
        "warning" { "⚠"; break }
        "info" { "ℹ"; break }
        default { "•"; break }
    }
    
    $markdownLine = "- [$timestamp] $statusIcon "
    
    if ($IsRemoved) {
        $markdownLine += "~~$Message~~"
    } else {
        $markdownLine += $Message
    }
    
    Add-Content -Path "Progress.MD" -Value $markdownLine
}

# Function to update Excel column names
function Update-ExcelColumnNames {
    param (
        [string]$ExcelPath,
        [string]$NewExcelPath,
        [hashtable]$ColumnMapping
    )
    
    try {
        $excel = Open-ExcelPackage -Path $ExcelPath
        $worksheet = $excel.Workbook.Worksheets[1]
        
        # Get all column headers
        $headers = @()
        $col = 1
        while ($col -le $worksheet.Dimension.End.Column) {
            $headers += $worksheet.Cells[1, $col].Text
            $col++
        }
        
        # Check if ImportStatus column already exists
        $importStatusExists = $headers -contains "ImportStatus"
        
        # Update column headers based on mapping
        $col = 1
        foreach ($header in $headers) {
            if ($ColumnMapping.ContainsKey($header)) {
                $worksheet.Cells[1, $col].Value = $ColumnMapping[$header]
            }
            $col++
        }
        
        # Add ImportStatus column only if it doesn't exist
        if (-not $importStatusExists) {
            $newCol = $worksheet.Dimension.End.Column + 1
            $worksheet.Cells[1, $newCol].Value = "ImportStatus"
            
            # Initialize all rows with 0
            2..$worksheet.Dimension.End.Row | ForEach-Object {
                $worksheet.Cells[$_, $newCol].Value = 0
            }
        }
        
        # Save as new file
        $excel | Close-ExcelPackage -NoSave:$false
        Copy-Item -Path $ExcelPath -Destination $NewExcelPath -Force
        
        Write-Host "Excel column names updated and new file created"
        Write-ProgressToMD "Excel column names updated and new file created"
        
    }
    catch {
        Write-Host "Error updating Excel: $($_.Exception.Message)" -ForegroundColor Red
        Write-ProgressToMD "Error updating Excel: $($_.Exception.Message)"
        throw
    }
}

# Function to update ImportStatus in Excel
function Update-ImportStatus {
    param (
        [string]$ExcelPath,
        [int]$RowNumber
    )
    
    try {
        $excel = Open-ExcelPackage -Path $ExcelPath
        $worksheet = $excel.Workbook.Worksheets[1]
        
        # Find ImportStatus column
        $importStatusColumn = 1
        $importStatusFound = $false
        
        while ($importStatusColumn -le $worksheet.Dimension.End.Column) {
            if ($worksheet.Cells[1, $importStatusColumn].Text -eq "ImportStatus") {
                $importStatusFound = $true
                break
            }
            $importStatusColumn++
        }
        
        if ($importStatusFound) {
            # Update ImportStatus value to 1
            # RowNumber 1'den başlıyor (ilk veri satırı) ve biz bunu Excel'de 2. satıra yazmalıyız (1. satır başlık)
            $actualRow = $RowNumber + 1
            $worksheet.Cells[$actualRow, $importStatusColumn].Value = 1
            
            # Save Excel and close
            $excel | Close-ExcelPackage -NoSave:$false
            Write-Host "Excel updated: $ExcelPath - Row: $RowNumber"
        }
        else {
            Write-Host "ImportStatus column not found: $ExcelPath" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error updating Excel: $ExcelPath - Row: $RowNumber - Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to check if Excel files are in use
function Test-ExcelFilesInUse {
    param (
        [string]$SourceExcel,
        [string]$UpdatedExcel
    )
    
    $files = @($SourceExcel)
    if (Test-Path $UpdatedExcel) {
        $files += $UpdatedExcel
    }
    
    $filesInUse = @()
    
    foreach ($file in $files) {
        try {
            $stream = [System.IO.File]::Open($file, 'Open', 'Read', 'None')
            $stream.Close()
            $stream.Dispose()
        }
        catch {
            $filesInUse += [System.IO.Path]::GetFileName($file)
        }
    }
    
    return $filesInUse
}

# Function to wait until Excel files are closed
function Wait-ForExcelFiles {
    param (
        [string]$SourceExcel,
        [string]$UpdatedExcel
    )
    
    $filesInUse = Test-ExcelFilesInUse -SourceExcel $SourceExcel -UpdatedExcel $UpdatedExcel
    
    while ($filesInUse.Count -gt 0) {
        Write-Host "Please close the following Excel files:" -ForegroundColor Yellow
        $filesInUse | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }
        Write-Host "Press Enter to continue after closing the files..." -ForegroundColor Yellow
        
        $null = Read-Host
        $filesInUse = Test-ExcelFilesInUse -SourceExcel $SourceExcel -UpdatedExcel $UpdatedExcel
    }
}

# Function to display mapping table
function Write-MappingTable {
    param (
        [hashtable]$Mapping,
        [object[]]$SharePointColumns
    )
    
    Write-Section "Column Mapping Preview"
    
    $format = "│ {0,-30} │ {1,-30} │ {2,-30} │"
    $line = "─" * 97
    $header = "┌" + ("─" * 32) + "┬" + ("─" * 32) + "┬" + ("─" * 32) + "┐"
    $separator = "├" + ("─" * 32) + "┼" + ("─" * 32) + "┼" + ("─" * 32) + "┤"
    $footer = "└" + ("─" * 32) + "┴" + ("─" * 32) + "┴" + ("─" * 32) + "┘"
    
    Write-Host
    Write-Host $header -ForegroundColor DarkCyan
    Write-Host ($format -f "Excel Column", "SharePoint Title", "SharePoint Internal") -ForegroundColor Cyan
    Write-Host $separator -ForegroundColor DarkCyan
    
    foreach ($map in $Mapping.GetEnumerator() | Sort-Object Key) {
        $spColumn = $SharePointColumns | Where-Object { $_.InternalName -eq $map.Value }
        if ($spColumn) {
            Write-Host ($format -f $map.Key, $spColumn.Title, $map.Value) -ForegroundColor White
        } else {
            Write-Host ($format -f $map.Key, "NOT FOUND", "NOT FOUND") -ForegroundColor Red
        }
    }
    
    Write-Host $footer -ForegroundColor DarkCyan
    Write-Host
    
    # Display statistics
    $totalColumns = $Mapping.Count
    $matchedColumns = ($SharePointColumns | Where-Object { $Mapping.ContainsValue($_.InternalName) }).Count
    
    Write-Host "Statistics:" -ForegroundColor DarkCyan
    Write-Host "• Total Excel Columns: $totalColumns" -ForegroundColor White
    Write-Host "• Matched Columns: $matchedColumns" -ForegroundColor Green
    Write-Host "• Unmatched Columns: $($totalColumns - $matchedColumns)" -ForegroundColor $(if ($totalColumns - $matchedColumns -gt 0) { "Red" } else { "Green" })
    Write-Host
}

# Load environment variables
$envPath = Join-Path $PSScriptRoot ".env"
Get-Content $envPath | ForEach-Object {
    if ($_ -match '^([^=]+)=(.*)$') {
        $key = $matches[1]
        $value = $matches[2].Trim('"')
        Set-Item -Path "env:$key" -Value $value
    }
}

# Ensure required modules are installed
Ensure-ModulesInstalled -ModuleNames $requiredModules

# Initialize Progress.MD if it doesn't exist
if (!(Test-Path "Progress.MD")) {
    Set-Content -Path "Progress.MD" -Value "# Excel to SharePoint Import Progress`n"
}

Write-Header "Excel to SharePoint Import Tool"
Write-Status "Starting import process..." -Status "info" -Color Cyan
Write-ProgressToMD "Starting import process..."

try {
    # Check Excel file existence
    $excelPath = Join-Path $PSScriptRoot $env:EXCEL_FILE
    if (!(Test-Path $excelPath)) {
        throw "Excel file not found: $excelPath"
    }
    
    # Wait until Excel files are closed
    Wait-ForExcelFiles -SourceExcel $excelPath -UpdatedExcel (Join-Path $PSScriptRoot $env:EXCEL_FILE_UPDATED)
    
    # Connect to SharePoint
    Write-Section "SharePoint Connection"
    Write-Status "Connecting to SharePoint site..." -Status "info" -Color Yellow
    Write-ProgressToMD "Connecting to SharePoint site: $env:SHAREPOINT_URL"
    
    Connect-PnPOnline -Url $env:SHAREPOINT_URL -UseWebLogin -WarningAction SilentlyContinue
    
    # Get SharePoint list columns
    Write-Status "Getting SharePoint list columns..." -Status "info" -Color Yellow
    Write-ProgressToMD "Getting SharePoint list columns: $env:LIST_NAME"
    
    $list = Get-PnPList -Identity $env:LIST_NAME
    $listColumns = Get-PnPField -List $list | Where-Object { -not $_.Hidden -and -not $_.ReadOnly }
    
    # Read Excel file
    Write-Section "Excel Processing"
    Write-Status "Reading Excel file..." -Status "info" -Color Yellow
    Write-ProgressToMD "Reading Excel file: $env:EXCEL_FILE"
    
    $excelData = Import-Excel -Path $excelPath
    
    # Get actual Excel columns and show debug info
    $excelColumns = $excelData[0].PSObject.Properties.Name
    Write-Status "Found Excel columns:" -Status "info" -Color Cyan
    $excelColumns | ForEach-Object { Write-Status "  • $_" -Status "info" -Color White }
    
    # Show SharePoint columns for debugging
    Write-Status "`nFound SharePoint columns:" -Status "info" -Color Cyan
    $listColumns | ForEach-Object { Write-Status "  • $($_.Title) [$($_.InternalName)]" -Status "info" -Color White }
    
    # Create column mapping only for existing Excel columns
    $columnMapping = @{}
    foreach ($excelColumn in $excelColumns) {
        $matchingColumn = $listColumns | Where-Object { 
            $_.Title -eq $excelColumn -or 
            $_.InternalName -eq $excelColumn -or
            $_.Title -eq $excelColumn.Trim()
        }
        if ($matchingColumn) {
            $columnMapping[$excelColumn] = $matchingColumn.InternalName
        }
    }
    
    # Display mapping table
    Write-MappingTable -Mapping $columnMapping -SharePointColumns $listColumns
    
    if ($columnMapping.Count -eq 0) {
        Write-Status "No column mappings found! Please check column names." -Status "error" -Color Red
        Write-ProgressToMD "No column mappings found. Please check column names." -Status "error"
        throw "No column mappings found. Excel columns might not match SharePoint columns."
    }
    
    $totalRows = $excelData.Count
    Write-Status "Found $totalRows rows to process" -Status "info" -Color Green
    Write-ProgressToMD "Found $totalRows rows to process"
    
    # Import data to SharePoint
    Write-Section "Data Import"
    $successCount = 0
    $failureCount = 0
    
    for ($i = 0; $i -lt $excelData.Count; $i++) {
        $row = $excelData[$i]
        
        # Skip if already imported
        if ($row.ImportStatus -eq 1) {
            Write-Status "Row $($i + 1) already imported, skipping..." -Status "info" -Color DarkGray
            Write-ProgressToMD "Row $($i + 1) already imported, skipping"
            continue
        }
        
        try {
            $itemHash = @{}
            foreach ($column in $listColumns) {
                if ($null -ne $row.$($column.Title)) {
                    $itemHash[$column.InternalName] = $row.$($column.Title)
                }
            }
            
            Add-PnPListItem -List $env:LIST_NAME -Values $itemHash | Out-Null
            
            # Update ImportStatus in Excel
            Update-ImportStatus -ExcelPath $excelPath -RowNumber ($i + 1)
            
            $successCount++
            Write-Status "Row $($i + 1) imported successfully" -Status "success" -Color Green
            Write-ProgressToMD "Row $($i + 1) imported successfully" -Status "success"
        }
        catch {
            $failureCount++
            Write-Status "Error importing row $($i + 1): $($_.Exception.Message)" -Status "error" -Color Red
            Write-ProgressToMD "Error importing row $($i + 1): $($_.Exception.Message)" -Status "error"
        }
    }
    
    # Summary
    Write-Section "Import Summary"
    Write-Status "Import completed!" -Status "success" -Color Green
    Write-Status "Successful: $successCount" -Status "success" -Color Green
    Write-Status "Failed: $failureCount" -Status "error" -Color Red
    Write-Status "Total: $totalRows" -Status "info" -Color Cyan
    
    Write-ProgressToMD "Import completed. Successful: $successCount, Failed: $failureCount, Total: $totalRows" -Status "success"
}
catch {
    $errorMessage = "Error: $($_.Exception.Message)"
    Write-Status $errorMessage -Status "error" -Color Red
    Write-ProgressToMD $errorMessage -Status "error"
}
finally {
    try {
        # Try to get the connection state without throwing an error
        $connection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($null -ne $connection -and $connection.Url) {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-Status "SharePoint connection closed" -Status "info" -Color Yellow
            Write-ProgressToMD "SharePoint connection closed"
        }
    }
    catch {
        # Connection might already be closed, just continue
        Write-Status "SharePoint connection already closed" -Status "info" -Color Yellow
        Write-ProgressToMD "SharePoint connection already closed"
    }
}

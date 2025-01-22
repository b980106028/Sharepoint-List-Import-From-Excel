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
    
    Write-Host "`n┌─────────────────────────────────────┐" -ForegroundColor DarkCyan
    Write-Host "│ $Title" -ForegroundColor DarkCyan
    Write-Host "└─────────────────────────────────────┘" -ForegroundColor DarkCyan
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
        
        Write-Host "Excel sütun başlıkları başarıyla güncellendi ve yeni dosya oluşturuldu"
        Write-ProgressToMD "Excel sütun başlıkları başarıyla güncellendi ve yeni dosya oluşturuldu"
        
    }
    catch {
        Write-Host "Excel güncellenirken hata oluştu: $($_.Exception.Message)" -ForegroundColor Red
        Write-ProgressToMD "Excel güncellenirken hata oluştu: $($_.Exception.Message)"
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
        
        # ImportStatus kolonunu bul
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
            # ImportStatus değerini 1 olarak güncelle
            # RowNumber 1'den başlıyor (ilk veri satırı) ve biz bunu Excel'de 2. satıra yazmalıyız (1. satır başlık)
            $actualRow = $RowNumber + 1
            $worksheet.Cells[$actualRow, $importStatusColumn].Value = 1
            
            # Excel'i kaydet ve kapat
            $excel | Close-ExcelPackage -NoSave:$false
            Write-Host "Excel güncellendi: $ExcelPath - Satır: $RowNumber"
        }
        else {
            Write-Host "ImportStatus kolonu bulunamadı: $ExcelPath" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Excel güncellenirken hata oluştu: $ExcelPath - Satır: $RowNumber - Hata: $($_.Exception.Message)" -ForegroundColor Red
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
        Write-Host "Lütfen aşağıdaki Excel dosyalarını kapatın:" -ForegroundColor Yellow
        $filesInUse | ForEach-Object { Write-Host "- $_" -ForegroundColor Yellow }
        Write-Host "Dosyaları kapattıktan sonra devam etmek için Enter'a basın..." -ForegroundColor Yellow
        
        $null = Read-Host
        $filesInUse = Test-ExcelFilesInUse -SourceExcel $SourceExcel -UpdatedExcel $UpdatedExcel
    }
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
    Write-Status "Connecting to SharePoint..." -Status "info" -Color Yellow
    Write-ProgressToMD "SharePoint sitesine bağlanılıyor: $env:SHAREPOINT_URL"
    
    Connect-PnPOnline -Url $env:SHAREPOINT_URL -UseWebLogin
    
    # Get SharePoint list columns
    Write-Status "Getting SharePoint list columns..." -Status "info" -Color Yellow
    Write-ProgressToMD "SharePoint liste sütunları getiriliyor: $env:LIST_NAME"
    
    $list = Get-PnPList -Identity $env:LIST_NAME
    $listColumns = Get-PnPField -List $list | Where-Object { -not $_.Hidden -and -not $_.ReadOnly }
    
    # Read Excel file
    Write-Section "Excel Processing"
    Write-Status "Reading Excel file..." -Status "info" -Color Yellow
    Write-ProgressToMD "Excel dosyası okunuyor: $env:EXCEL_FILE"
    
    $excelData = Import-Excel -Path $excelPath
    
    # Create column mapping
    $columnMapping = @{}
    foreach ($column in $listColumns) {
        $columnMapping[$column.Title] = $column.InternalName
    }
    
    # Update Excel headers
    Write-Status "Do you want to update Excel column headers with SharePoint internal names? (E/H)" -Status "warning" -Color Yellow
    $confirmation = Read-Host
    
    if ($confirmation -eq "E") {
        $newExcelPath = Join-Path $PSScriptRoot $env:EXCEL_FILE_UPDATED
        Write-Status "Updating Excel headers..." -Status "info" -Color Yellow
        Write-ProgressToMD "Excel sütun başlıkları SharePoint InternalName değerleri ile güncelleniyor ve yeni dosya oluşturuluyor: $($env:EXCEL_FILE_UPDATED)"
        Update-ExcelColumnNames -ExcelPath $excelPath -NewExcelPath $newExcelPath -ColumnMapping $columnMapping
        
        # Read updated Excel file
        Write-Status "Reading updated Excel file..." -Status "info" -Color Yellow
        Write-ProgressToMD "Güncellenmiş Excel dosyası okunuyor: $($env:EXCEL_FILE_UPDATED)"
        $excelData = Import-Excel -Path $newExcelPath
    }
    
    $totalRows = $excelData.Count
    Write-Status "Found $totalRows rows to process" -Status "info" -Color Green
    Write-ProgressToMD "Toplam $totalRows satır bulundu"
    
    # Import data to SharePoint
    Write-Section "Data Import"
    $successCount = 0
    $failureCount = 0
    
    for ($i = 0; $i -lt $excelData.Count; $i++) {
        $row = $excelData[$i]
        
        # Skip if already imported
        if ($row.ImportStatus -eq 1) {
            Write-Status "Row $($i + 1) already imported, skipping..." -Status "info" -Color DarkGray
            Write-ProgressToMD "$($i + 1). satır daha önce aktarılmış, atlandı"
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
            
            # Update ImportStatus in both Excel files
            Update-ImportStatus -ExcelPath $excelPath -RowNumber ($i + 1)
            if ($confirmation -eq "E") {
                Update-ImportStatus -ExcelPath $newExcelPath -RowNumber ($i + 1)
            }
            
            $successCount++
            Write-Status "Row $($i + 1) imported successfully" -Status "success" -Color Green
            Write-ProgressToMD "$($i + 1). satır başarıyla aktarıldı" -Status "success"
        }
        catch {
            $failureCount++
            Write-Status "Error importing row $($i + 1): $($_.Exception.Message)" -Status "error" -Color Red
            Write-ProgressToMD "$($i + 1). satır aktarılamadı: $($_.Exception.Message)" -Status "error"
        }
    }
    
    # Summary
    Write-Section "Import Summary"
    Write-Status "Import completed!" -Status "success" -Color Green
    Write-Status "Successful: $successCount" -Status "success" -Color Green
    Write-Status "Failed: $failureCount" -Status "error" -Color Red
    Write-Status "Total: $totalRows" -Status "info" -Color Cyan
    
    Write-ProgressToMD "Aktarım tamamlandı. Başarılı: $successCount, Başarısız: $failureCount, Toplam: $totalRows" -Status "success"
    
    # Disconnect from SharePoint
    Disconnect-PnPOnline
    Write-Status "SharePoint connection closed" -Status "info" -Color Yellow
    Write-ProgressToMD "SharePoint bağlantısı kapatıldı"
}
catch {
    $errorMessage = "Error: $($_.Exception.Message)"
    Write-Status $errorMessage -Status "error" -Color Red
    Write-ProgressToMD $errorMessage -Status "error"
}
finally {
    if (Get-PnPConnection) {
        Disconnect-PnPOnline
        Write-Status "SharePoint connection closed" -Status "info" -Color Yellow
        Write-ProgressToMD "SharePoint bağlantısı kapatıldı"
    }
}

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

# Function to write progress to MD file
function Write-ProgressToMD {
    param (
        [string]$Message,
        [bool]$IsRemoved = $false
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $markdownLine = "- [$timestamp] "
    
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

Write-ProgressToMD "Starting import process..."

try {
    # Check Excel file existence
    $excelPath = Join-Path $PSScriptRoot $env:EXCEL_FILE
    if (!(Test-Path $excelPath)) {
        throw "Excel dosyası bulunamadı: $excelPath"
    }
    
    # Wait until Excel files are closed
    Wait-ForExcelFiles -SourceExcel $excelPath -UpdatedExcel (Join-Path $PSScriptRoot $env:EXCEL_FILE_UPDATED)

    # Connect to SharePoint
    Write-Host "SharePoint'e bağlanılıyor..."
    Write-ProgressToMD "SharePoint sitesine bağlanılıyor: $env:SHAREPOINT_URL"
    Connect-PnPOnline -Url $env:SHAREPOINT_URL -UseWebLogin

    # SharePoint liste sütunlarını al
    Write-Host "`nSharePoint liste sütunları getiriliyor..."
    Write-ProgressToMD "SharePoint liste sütunları getiriliyor: $env:LIST_NAME"
    $list = Get-PnPList -Identity $env:LIST_NAME
    $fields = Get-PnPField -List $list | Where-Object { -not $_.Hidden -and $_.InternalName -notlike "Computed_*" -and $_.InternalName -notlike "_*" }
    
    Write-Host "`nListe Sütunları:"
    Write-Host "----------------"
    $fields | Select-Object Title, InternalName | Format-Table -AutoSize

    # Excel'i oku ve sütunları göster
    Write-Host "`nExcel dosyası okunuyor: $excelPath"
    Write-ProgressToMD "Excel dosyası okunuyor: $env:EXCEL_FILE"
    
    $excelData = Import-Excel -Path $excelPath
    if ($null -eq $excelData -or $excelData.Count -eq 0) {
        throw "Excel dosyası boş veya okunamadı"
    }

    # Excel sütunlarını göster
    Write-Host "`nExcel Sütunları:"
    Write-Host "---------------"
    $excelColumns = $excelData[0].PSObject.Properties.Name
    $excelColumns | ForEach-Object { Write-Host $_ }

    # Eşleştirmeleri oluştur
    $columnMapping = @{}
    foreach ($excelCol in $excelColumns) {
        $matchingField = $fields | Where-Object { $_.Title -eq $excelCol }
        if ($matchingField) {
            $columnMapping[$excelCol] = $matchingField.InternalName
        }
    }

    # Eşleştirmeleri göster ve onay iste
    Write-Host "`nÖnerilen Eşleştirmeler:"
    Write-Host "----------------------"
    foreach ($mapping in $columnMapping.GetEnumerator()) {
        Write-Host "Excel Kolonu: $($mapping.Key)"
        Write-Host "SharePoint Title: $($fields | Where-Object { $_.InternalName -eq $mapping.Value } | Select-Object -ExpandProperty Title)"
        Write-Host "SharePoint InternalName: $($mapping.Value)"
        Write-Host "----------------------------------------"
    }

    $confirmation = Read-Host "`nYukarıdaki eşleştirmeler doğru mu? Excel sütun başlıkları SharePoint InternalName değerleri ile değiştirilecek. (E/H)"
    
    if ($confirmation -eq "E") {
        $newExcelPath = Join-Path $PSScriptRoot $env:EXCEL_FILE_UPDATED
        Write-Host "`nExcel sütun başlıkları güncelleniyor ve yeni dosya oluşturuluyor: $($env:EXCEL_FILE_UPDATED)"
        Write-ProgressToMD "Excel sütun başlıkları SharePoint InternalName değerleri ile güncelleniyor ve yeni dosya oluşturuluyor: $($env:EXCEL_FILE_UPDATED)"
        Update-ExcelColumnNames -ExcelPath $excelPath -NewExcelPath $newExcelPath -ColumnMapping $columnMapping
        Write-Host "Excel sütun başlıkları başarıyla güncellendi ve yeni dosya oluşturuldu!"
        Write-ProgressToMD "Excel sütun başlıkları başarıyla güncellendi ve yeni dosya oluşturuldu"

        # Yeni Excel dosyasını oku
        Write-Host "`nGüncellenmiş Excel dosyası okunuyor: $newExcelPath"
        Write-ProgressToMD "Güncellenmiş Excel dosyası okunuyor: $($env:EXCEL_FILE_UPDATED)"
        
        $excelData = Import-Excel -Path $newExcelPath
        $totalRows = $excelData.Count
        Write-Host "Toplam aktarılacak satır: $totalRows"
        Write-ProgressToMD "Toplam $totalRows satır bulundu"

        Write-Host "`nGüncellenmiş Excel Sütunları:"
        Write-Host "-------------------------"
        $excelColumns = $excelData[0].PSObject.Properties.Name
        $excelColumns | ForEach-Object { Write-Host $_ }

        # Import each row
        $successCount = 0
        $failureCount = 0
        
        for ($i = 0; $i -lt $totalRows; $i++) {
            $row = $excelData[$i]
            
            # ImportStatus 1 ise bu satırı atla
            if ($row.ImportStatus -eq 1) {
                Write-Host "Satır $($i + 1) daha önce aktarılmış, atlanıyor..."
                Write-ProgressToMD "$($i + 1). satır daha önce aktarılmış, atlandı"
                continue
            }
            
            $progress = [math]::Round(($i + 1) / $totalRows * 100, 2)
            Write-Progress -Activity "SharePoint'e veriler aktarılıyor" -Status "%$progress Tamamlandı" -PercentComplete $progress
            Write-Host "İşlenen satır $($i + 1) / $totalRows"
            
            try {
                # SharePoint'e gönderilecek özellikleri hazırla (ImportStatus hariç)
                $itemProperties = @{}
                $row.PSObject.Properties | Where-Object { $_.Name -ne "ImportStatus" } | ForEach-Object {
                    if ($null -ne $_.Value) {
                        $itemProperties[$_.Name] = $_.Value
                    }
                }
                
                # Add item to SharePoint list
                Add-PnPListItem -List $env:LIST_NAME -Values $itemProperties
                $successCount++
                Write-ProgressToMD "$($i + 1). satır başarıyla aktarıldı"
                
                # Her iki Excel dosyasında da ImportStatus'u güncelle
                Update-ImportStatus -ExcelPath $excelPath -RowNumber ($i + 1)
                Update-ImportStatus -ExcelPath $newExcelPath -RowNumber ($i + 1)
            }
            catch {
                $failureCount++
                Write-ProgressToMD "$($i + 1). satır aktarılamadı: $($_.Exception.Message)"
                Write-Host "Hata - $($i + 1). satır: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Write summary
        $summary = "Aktarım tamamlandı. Başarılı: $successCount, Başarısız: $failureCount, Toplam: $totalRows"
        Write-Host $summary
        Write-ProgressToMD $summary
    }
    else {
        Write-Host "İşlem iptal edildi."
        Write-ProgressToMD "Sütun eşleştirme işlemi kullanıcı tarafından iptal edildi"
    }
}
catch {
    $errorMessage = "Hata: $($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    Write-ProgressToMD $errorMessage
}
finally {
    # Disconnect from SharePoint
    try {
        Disconnect-PnPOnline
        Write-ProgressToMD "SharePoint bağlantısı kapatıldı"
    } catch {
        Write-ProgressToMD "SharePoint bağlantısı zaten kapalı"
    }
}

# Excel'den SharePoint'e Veri Aktarım Scripti

Bu PowerShell scripti, Excel dosyasından SharePoint listesine veri aktarımı yapar.

## Özellikler

- Excel sütun başlıklarını SharePoint kolon isimleri ile eşleştirir
- Daha önce aktarılan satırları atlar (ImportStatus kolonu ile kontrol)
- İlerleme durumunu Progress.MD dosyasında tutar
- Excel dosyalarının açık olup olmadığını kontrol eder

## Kurulum

1. `.env.example` dosyasını `.env` olarak kopyalayın
2. `.env` dosyasındaki değişkenleri kendi ortamınıza göre güncelleyin:
   - SHAREPOINT_URL: SharePoint site URL'i
   - LIST_NAME: SharePoint liste adı
   - EXCEL_FILE: Kaynak Excel dosyası
   - EXCEL_FILE_UPDATED: Güncellenmiş Excel dosyası

## Kullanım

```powershell
.\Import-ExcelToSharePoint.ps1
```

## Gereksinimler

- PowerShell 5.1 veya üzeri
- PnP.PowerShell modülü
- ImportExcel modülü

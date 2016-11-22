$ExcelObject = New-Object -ComObject Excel.Application
$ExcelObject.visible = $false
$ExcelObject.DisplayAlerts = $false
$date = Get-Date -Format "yyyyMMddHHss"
$strPath1 = "o:\UserCert\Active_Users_$date.xlsx"
if (Test-Path $strPath1) {
  #Open the document  
  $ActiveWorkbook = $ExcelObject.WorkBooks.Open($strPath1)
  $ActiveWorksheet = $ActiveWorkbook.Worksheets.Item(1)
}

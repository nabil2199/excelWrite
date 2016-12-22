
$file = "C:\Repositories\excelWrite\ExcelFile.xlsx"
$sheetName = "Sheet1"
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false
#Count max row
$rowMax = ($sheet.UsedRange.Rows).count 
#Declare the starting positions
$rowName,$colName = 1,1
$rowAge,$colAge = 1,2
$rowCity,$colCity = 1,3
#loop to get values and store it
for ($i=1; $i -le $rowMax-1; $i++)
{
$name = $sheet.Cells.Item($rowName+$i,$colName).text 
$age = $sheet.Cells.Item($rowAge+$i,$colAge).text 
$city = $sheet.Cells.Item($rowCity+$i,$colCity).text 

Write-Host ("My Name is: "+$name)
Write-Host ("My Age is: "+$age)
Write-Host ("I live in: "+$city)
}
#close excel file
$objExcel.quit()
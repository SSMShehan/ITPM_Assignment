$csvPath = "$PWD\data.csv"
$excelPath = "$PWD\IT23836068_Assignment 1 - Test cases.xlsx"

if (Test-Path $excelPath) { Remove-Item $excelPath }

$data = Import-Csv -Path $csvPath -Delimiter "|" -Encoding UTF8

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)
$sheet.Name = " Test cases"

# Write headers
$headers = @("TC_ID", "Input length type", "Input", "Expected output", "Actual output", "Status", "Singlish input types covered", "Evidence or rationale for the input type covered")
for ($c = 1; $c -le 8; $c++) {
    $sheet.Cells.Item(1, $c) = $headers[$c-1]
}

# Write data
$row = 2
foreach ($item in $data) {
    $sheet.Cells.Item($row, 1) = $item."TC_ID"
    $sheet.Cells.Item($row, 2) = $item."Input length type"
    $sheet.Cells.Item($row, 3) = $item."Input"
    $sheet.Cells.Item($row, 4) = $item."Expected output"
    $sheet.Cells.Item($row, 5) = $item."Actual output"
    $sheet.Cells.Item($row, 6) = $item."Status"
    $sheet.Cells.Item($row, 7) = $item."Singlish input types covered"
    $sheet.Cells.Item($row, 8) = $item."Evidence or rationale for the input type covered"
    $row++
}

# Formatting
$headerRange = $sheet.Range("A1:H1")
$headerRange.Font.Bold = $true
$headerRange.Font.Color = 0x000000 # Black font
$headerRange.Interior.Color = 0xD9D9D9 # Light Gray
$headerRange.HorizontalAlignment = -4108 # xlCenter
$headerRange.VerticalAlignment = -4108 # xlCenter

$fullRange = $sheet.Range("A1:H51")

# Borders
$fullRange.Borders.LineStyle = 1 # xlContinuous
$fullRange.Borders.Weight = 2 # xlThin

# Alignment
$fullRange.VerticalAlignment = -4108 # xlCenter
$sheet.Range("C2:E51").HorizontalAlignment = -4131 # xlLeft
$sheet.Range("G2:H51").HorizontalAlignment = -4131 # xlLeft
$sheet.Range("A2:B51").HorizontalAlignment = -4108 # xlCenter
$sheet.Range("F2:F51").HorizontalAlignment = -4108 # xlCenter

# Wrap Text
$fullRange.WrapText = $true

# Column Widths
$sheet.Columns.Item(1).ColumnWidth = 12
$sheet.Columns.Item(2).ColumnWidth = 15
$sheet.Columns.Item(3).ColumnWidth = 45
$sheet.Columns.Item(4).ColumnWidth = 45
$sheet.Columns.Item(5).ColumnWidth = 35
$sheet.Columns.Item(6).ColumnWidth = 10
$sheet.Columns.Item(7).ColumnWidth = 35
$sheet.Columns.Item(8).ColumnWidth = 60

# Freeze Top Row
$excel.ActiveWindow.SplitRow = 1
$excel.ActiveWindow.FreezePanes = $true

# AutoFit row heights
$fullRange.EntireRow.AutoFit()

$workbook.SaveAs($excelPath)
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "Excel sheet successfully generated from CSV with proper formatting."

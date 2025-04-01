param (
    [string]$excelPath,
    [string]$sheetName,
    [string]$csvPath
)

if (-not $excelPath -or -not $sheetName) {
    Write-Host "Forma de Uso: .\ConvertExcelToCSV.ps1 <ruta_excel> <nombre_hoja> [ruta_csv]"
    exit 1
}

if (-not $csvPath) {
    $csvPath = [System.IO.Path]::ChangeExtension($excelPath, ".csv")
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    $sheet = $workbook.Sheets.Item($sheetName)

    # Exporta como CSV
    $sheet.SaveAs($csvPath, 6)  # 6 es el código para CSV

    Write-Host "Conversión completada: $csvPath"
}
catch {
    Write-Host "Error: No se pudo abrir el archivo o la hoja '$sheetName' no existe."
}
finally {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
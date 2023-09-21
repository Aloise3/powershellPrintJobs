

$sourceFolder = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest"

$pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf

Write-Host "PDF file found: '$pdfFiles'"
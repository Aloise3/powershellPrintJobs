# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1

function Start-SendPDFtoPrint {
    param (
        [string]$SmtpServer,
        [string]$ExcelFilePath, #Bruges
        [string]$senderEmail = "PrintJobs", 
        [string]$recipientEmail, #Bruges
        [string]$sourceFolder,
        [string]$destinationFolder,
        [string]$printerName, 
        [int]$runspecificstuff = 1
    )



    # Get a list of PDF files in the source folder
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf

    
    $timestamp = Get-Date -Format "dd/MM/yyyy kl. hh:mm" #tidspunkt
    # Check if any PDF files are found
    if ($pdfFiles.Count -gt 0) {
        # Loop through each PDF file found
        foreach ($pdfFile in $pdfFiles) {
            
            if ($runspecificstuff -eq 1) {

                if ($ExcelFilePath) {
                    # Laver en excel-fil
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false
                    $workbook = $excel.Workbooks.Open($ExcelFilePath)  # Bruger prædefineret sti
                    $worksheet = $workbook.Worksheets.Item(1) # Første side

                    # Kolonneoverskrifter
                    $worksheet.Cells.Item(1, 1).Value2 = "Fil"
                    $worksheet.Cells.Item(1, 2).Value2 = "Printet tidspunkt"
                    $worksheet.Cells.Item(1, 3).Value2 = "Bruger"

                    # Append data to Excel
                    $row = $worksheet.UsedRange.Rows.Count + 1
                    $worksheet.Cells.Item($row, 1).Value2 = $pdfFile
                    $worksheet.Cells.Item($row, 2).Value2 = $timestamp
                    $worksheet.Cells.Item($row, 3).Value2 = $Env:UserName

                    # Gemmer exceloversigt
                    $workbook.Save()
                    $excel.Quit()

                    }

                if ($printerName) {
                    # Print the PDF file to the specified printer
                    Start-Job -ScriptBlock {
                    param ($pdfFile, $printerName)
                    $pdfFile | Out-Printer -Name $printerName
                    } -ArgumentList $pdfFile.FullName, $printerName | Wait-Job | Receive-Job
            
                    }

                    else {
                        Start-Process -FilePath $pdfFile.FullName -Verb Print -PassThru | ForEach-Object {
                        # Vent på printeren
                        $_ | Wait-PrinterJob 
                        }
                    } 
            }
            # Flyt filen til arkiv
            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force
            
            # Send en mail
            Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Der er printet en fil - sendt til printer '$printerName'"  -Body "Attached is the PDF file that was printed." -Attachments $pdfFile.FullName
            
        }
    }


}

Export-ModuleMember -Function Start-SendPDFtoPrint



# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1

function Start-SendPDFtoPrint {
    param (
        [string]$SmtpServer,
        [string]$ExcelFilePath, #Bruges
        [string]$senderEmail = "PrintJobs", 
        [string]$recipientEmail, #Bruges
        [string]$sourceFolder,
        [string]$destinationFolder,
        [string]$printerName 
    )



    # Get a list of PDF files in the source folder
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf

    Write-Host "PDF fil fundet: '$pdfFiles'"
    $timestamp = Get-Date -Format "dd/MM/yyyy kl. hh:mm" #tidspunkt
    # Check if any PDF files are found
    if ($pdfFiles.Count -gt 0) {
        # Loop through each PDF file found
        foreach ($pdfFile in $pdfFiles) {
            # Move the PDF file to the destination folder
            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force
            Write-Host "Fil er rykket"

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
            
            # Send en mail
            Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Der er printet en fil - sendt til printer '$printerName'"  -Body "Attached is the PDF file that was printed." -Attachments $pdfFile.FullName
            
        }
    }




Export-ModuleMember -Function Start-SendPDFtoPrint



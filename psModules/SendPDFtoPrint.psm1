# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1

function Write-Log {
    param (
        [string]$name,
        [string]$time,
        [string]$user,
        [string]$logFilePath
    )

    if ($logFilePath -and $logFilePath -like "*.xlsx") {
                    
        # Laver en excel-fil
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($logFilePath)  # Bruger prædefineret sti
        $worksheet = $workbook.Worksheets.Item(1) # Første side

        # Kolonneoverskrifter
        $worksheet.Cells.Item(1, 1).Value2 = "Fil"
        $worksheet.Cells.Item(1, 2).Value2 = "Printet tidspunkt"
        $worksheet.Cells.Item(1, 3).Value2 = "Bruger"
        # Append data to Excel
        $row = $worksheet.UsedRange.Rows.Count + 1
        $worksheet.Cells.Item($row, 1).Value2 = $name
        $worksheet.Cells.Item($row, 2).Value2 = $time
        $worksheet.Cells.Item($row, 3).Value2 = $user
        $worksheet.UsedRange.EntireColumn.AutoFit()
        # Gemmer exceloversigt
        $workbook.Save()
        $excel.Quit() 
    }

}

function Start-execPrinter {
        param (
            [string]$printerName,
            [string]$pdfFile
        )

        if ($printerName) {
            # Send til specifik printer
            Start-Job -ScriptBlock {
            param ($pdfFile, $printerName)
            $pdfFile | Out-Printer -Name $printerName
            } -ArgumentList $pdfFile.FullName, $printerName | Wait-Job | Receive-Job
        } 
        else 
        {
            Start-Process -FilePath $pdfFile.FullName -Verb Print -PassThru | ForEach-Object {
            # Vent på printeren
            $_ | Wait-PrinterJob }
        } 
}




function Start-SendPDFtoPrint {
    param (
        [string]$SmtpServer,
        [string]$logPath, #Bruges
        [string]$senderEmail = "PrintJobs", 
        [string]$recipientEmail, #Bruges
        [string]$sourceFolder,
        [string]$destinationFolder,
        [string]$printerName, 
        [string]$user = $env:USERNAME,
        [int]$runspecificstuff = 1
    )



    # Få en liste over pdf filer
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf

    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #tidspunkt
    # Se om der er fundet pdf-filer
    if ($pdfFiles.Count -gt 0) {
        # loop gennem alle filer (hvis man hurtigt tilføjer flere)
        foreach ($pdfFile in $pdfFiles) {
            Write-Host "navn: '$pdfFiles'"
            Write-Host "tidspunkt: '$timestamp'"
            Write-Host "bruger: '$env:USERNAME'"
            
            
            if ($logPath) {
                Write-Log -name $pdfFile -time $timestamp -user $env:USERNAME -logFilePath $logPath
                }

            
            if ($runspecificstuff -eq 1) {
                Start-execPrinter -printerName $printerName -pdfFile $pdfFiles
            }
            # Flyt filen til arkiv

            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force
            

            <# # Send en mail
            try {
                Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Der er printet en fil - sendt til printer '$printerName'"  -Body "Attached is the PDF file that was printed." -Attachments $pdfFile.FullName
            } catch {
                if (-not $errvariable) {
                $errvariable = "Fejl: Mailafsendelse"
                }
                $body = "Mail for printjob med vedhæftning har fejlet pga. '$errvariable'"
            }
            finally {
                Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Printjob har fejlet'"  -Body $body 
            } #>
        
        
        }
    }
}

Export-ModuleMember -Function Start-SendPDFtoPrint



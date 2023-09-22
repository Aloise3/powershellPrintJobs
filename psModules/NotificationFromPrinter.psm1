# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1

function Start-PrintJobMonitor {
    param (
        [string]$SmtpServer,
        [string]$userName, #Bruges
        [string]$ExcelFilePath, #Bruges
        [string]$senderEmail = "PrintJobs", 
        [string]$recipientEmail, #Bruges
        [string]$TempPath = "C:/Temp",
        [string]$SourceIdentifier
    )

    # WMI Query til at tjekke Printjobs
    $query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'"

    
    # Register an event handler to execute when a print job is created
    Register-WmiEvent -Query $query -SourceIdentifier $SourceIdentifier -Action { #Tjekker om der er et nyt event
        $eventArgs = $Event.SourceEventArgs.NewEvent
        $printJob = $eventArgs.TargetInstance #Eventinformationer
        $fileName = $printJob.Document #navn på printjob
        $printerName = $printJob.HostPrintQueue #navn på printer
        $ownerName = $printJob.Owner
        $timestamp = Get-Date -Format "yyyy_MM_dd" #tidspunkt


        if ($ownerName = $userName) {
            
            
            # Kopierer dokumentet til en temporer folder

            try {
                #Udkommenteres da det er besværligt. Har også udkommenteret i mail
                # Laver et filnavn til vedhæftning i mail
                <#$attachmentFileName = $fileName+ "_"+ $timestamp + ".pdf"
                $destinationPath = Join-Path -Path $TempPath -ChildPath $attachmentFileName
                Copy-Item -Path $fileName -Destination $destinationPath #>

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
                    $worksheet.Cells.Item($row, 1).Value2 = $fileName
                    $worksheet.Cells.Item($row, 2).Value2 = $timestamp
                    $worksheet.Cells.Item($row, 3).Value2 = $ownerName

                    # Gemmer exceloversigt
                    $workbook.Save()
                    $excel.Quit()
                    }
                    
                # Sender en mail
                Send-MailMessage -From $senderEmail -To $recipientEmail -Subject "Print Job udført på '$filename'" -Body "Print job '$fileName' er blevet printet af '$ownerName' d. '$timestamp' i printer '$printerName' " -SmtpServer $SmtpServer #-Attachments $destinationPath

                }
            catch {
                Send-MailMessage -From $senderEmail -To $recipientEmail -Subject "Print Job fejl for $jobName!" -Body "Mail for printjob med vedhæftning har fejlet pga. at der er sat en forkert midlertidig sti." -SmtpServer $SmtpServer
                }
            finally {
                Write-Host "Jobbet $SourceIdentifier er sat op til modtager $recipientEmail"
                }

            }

    }
}

Export-ModuleMember -Function Start-PrintJobMonitor

function Stop-PrintJobMonitor {
    param (
        [string]$NameOfProces
    
    )
        Unregister-Event -SourceIdentifier $NameOfProces
}

Export-ModuleMember -Function Stop-PrintJobMonitor

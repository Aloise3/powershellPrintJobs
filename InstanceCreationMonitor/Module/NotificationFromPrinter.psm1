﻿# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1


function Write-Log {
    param (
        [string]$name,
        [string]$time,
        [string]$user,
        [string]$logFilePath
    )
    if ($logFilePath -like "*.xlsx") {
        Write-Host "Logfil er af .xlsx"            
        # Laver en excel-fil
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($logFilePath)  # Bruger prædefineret sti
        $worksheet = $workbook.Worksheets.Item(1) # Første side
        Write-Host "Logfil åbnet..."
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
    if ($logFilePath -like "*.txt") { #Hvis man foretrækker en tekstfil

        $logMessage = "$time - Filnavn: $name, Bruger: $user."
        $logMessage | Out-File -FilePath $logFilePath -Append
    }
}
Export-ModuleMember -Function Write-Log

function Start-PrintJobMonitor {
    param (
        [string]$SmtpServer,
        [string]$userName, #Bruges
        [string]$logpath, #Bruges
        [string]$senderEmail = "PrintJobs", 
        [string]$recipientEmail, #Bruges
        [string]$SourceIdentifier
    )

    # WMI Query til at tjekke Printjobs
    $query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PrintJob'"
    $Global:userName = $userName 
    
    # Registrer event der fyres af, når et nyt printerjob kommer
    Register-WmiEvent -Query $query -SourceIdentifier $SourceIdentifier -Action { #Tjekker om der er et nyt event
        $eventArgs = $Event.SourceEventArgs.NewEvent
        $printJob = $eventArgs.TargetInstance #Eventinformationer
        $fileName = $printJob.Document #navn på printjob
        $printerName = $printJob.HostPrintQueue #navn på printer
        $ownerName = $printJob.Owner
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #tidspunkt
        if ($ownerName -eq $Global:userName) {
          
            if ($logpath) {
                try {
                Write-Log -name $fileName -time $timestamp -user $Global:userName -logFilePath $logpath
                } catch {
                Write-Host "Error: $_"
                }                
            }

            if ($SmtpServer) {        
            # Sender en mail
                Send-MailMessage -From $senderEmail -To $recipientEmail -Subject "Print Job udført på '$filename'" -Body "Print job '$fileName' er blevet printet af '$ownerName' d. '$timestamp' i printer '$printerName'. Log findes på '$logpath' " -SmtpServer $SmtpServer #-Attachments $destinationPath
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

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
    if ($logFilePath -and $logFilePath -like "*.txt") { #Hvis man foretrækker en tekstfil

        $logMessage = "$time - Filnavn: $name, Bruger: $user."
        $logMessage | Out-File -FilePath $logFilePath -Append
    }
}


function Change-DefaultPrinter {
    param (
        [string]$printerName
    )

    $originalDefaultPrinter = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default=$true"
    $desiredPrinter = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Name='$printerName'"

    if ($desiredPrinter) {
        $desiredPrinter.SetDefaultPrinter()
    }

    return $originalDefaultPrinter
}

function Return-DefaultPrinter {
    param (
        [object]$originalDefaultPrinter
    )

    if ($originalDefaultPrinter) {
        $originalDefaultPrinter.SetDefaultPrinter()
    }
}

function Start-execPrinter {
        param (
            [string]$printerName,
            [string]$pdfFile
        )
              
        Start-Process -FilePath $pdfFile.FullName -Verb Print -PassThru | ForEach-Object {
        # Vent på printeren
        $_ | Wait-PrinterJob }
        
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
        [string]$user = $env:USERNAME
    )

    # Få en liste over pdf filer
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #tidspunkt

    if ($pdfFiles.Count -gt 0) {
        foreach ($pdfFile in $pdfFiles) {          
            
            if ($logPath) {
                try {
                    Write-Log -name $pdfFile.FullName -time $timestamp -user $env:USERNAME -logFilePath $logPath
                } catch {
                        $errvariable = "Fejl: Logning blev ikke gennemført"
                }
            }

            if ($printerName) {
                $originalDefaultPrinter = Change-DefaultPrinter -printerName $printerName
            }

            try {
                Start-execPrinter -printerName $printerName -pdfFile $pdfFile.FullName
            } catch {
                    $errvariable = "Fejl: Print blev ikke sendt"
            }

            if ($printerName) {
                Return-DefaultPrinter -originalDefaultPrinter $originalDefaultPrinter
            }
            
            if ($SmtpServer) {
                 # Send en mail
                try {
                    Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Der er printet en fil - sendt til printer '$printerName'"  -Body "Attached is the PDF file that was printed." -Attachments $pdfFile.FullName
                } catch {
                        if (-not $errvariable) {
                            $errvariable = "Fejl: Mailafsendelse"
                        }
                        $body = "Mail for printjob med vedhæftning har fejlet pga. '$errvariable'"
                        Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Printjob har fejlet'"  -Body $body
                }
            }
            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force #Rykker til arkiv
        }
    }
}

Export-ModuleMember -Function Start-SendPDFtoPrint



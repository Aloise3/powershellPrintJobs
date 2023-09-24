# C:\MyModules\PrintJobMonitor\PrintJobMonitor.psm1

function Write-Log {
    param (
        [string]$name,
        [string]$time,
        [string]$user,
        [string]$logFilePath
    )

    if ($logFilePath -like "*.xlsx") {
                    
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
    if ($logFilePath -like "*.txt") { #Hvis man foretrækker en tekstfil

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
        [string]$originalDefaultPrinter
    )
    $returnPrinter  = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Name='$originalDefaultPrinter'"
    $returnPrinter.SetDefaultPrinter()

}

function Start-execPrinter {
        param (
            [string]$pdfFile
        )
              
        Start-Process -FilePath $pdfFile.FullName -Verb Print -PassThru | ForEach-Object {
        # Vent på printeren
        $_ | Wait-PrinterJob }
        
}


function Start-SendPDFtoPrint {
    param (
        [Parameter(HelpMessage = "Optionel. Navn på SMTP-server. Bruges, hvis man vil sende en mail efter at filen er printet til fysisk printer.")]
        [string]$SmtpServer,
        [Parameter(HelpMessage = "Sti til en .xlsx eller .txt fil")]
        [ValidateScript({
            if (Test-Path $_ -PathType Leaf) {
                $extension = [System.IO.Path]::GetExtension($_)
                if ($extension -eq ".xlsx" -or $extension -eq ".txt") {
                    $true
                } else {
                    throw "Logfilen skal være af type .xlsx or .txt."
                }
            } else {
                throw "The specified path does not point to a valid file."
            }
        })][string]$logPath,
        [Parameter(HelpMessage = "Printernavn på en fysisk printer, som computeren er forbundet til")]
        [ValidateScript({
            $printer = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Name = '$_'"
            if ($printer -ne $null) {
                $true
            } else {
                throw "Printer '$_' does not exist or is not accessible on this computer."
            }
        })] [string]$printerName,
        [Parameter(HelpMessage = "Optionel. Navnet på afsender-email. Default: Printjobs")]
        [string]$senderEmail = "PrintJobs", 
        [Parameter(HelpMessage = "Optionel. Navnet på modtager-email. Kræver at SmtpServer er defineret.")]
        [string]$recipientEmail, 
        [Parameter(Mandatory=$true, HelpMessage = "Obligatorisk. Sti til folder, som indeholder den PDF-printede fil")]
        [string]$sourceFolder,
        [Parameter(Mandatory=$true, HelpMessage = "Obligatorisk. Sti til arkivering af den PDF-printede fil")]
        [string]$destinationFolder,
        [Parameter(HelpMessage = "Optionel. Bruger som logges for PDF-fil. Default: Brugeren på powershell-sessionen.")]
        [string]$user = $env:USERNAME
    )

    # Få en liste over pdf filer
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #tidspunkt

    if ($pdfFiles.Count -gt 0) {
        foreach ($pdfFile in $pdfFiles) {          
            
            if ($logPath) {
                try {
                    Write-Log -name $pdfFile.FullName -time $timestamp -user $user -logFilePath $logPath
                } catch {
                    $errvariable =  $_
                }
            }

            if ($printerName) { 
                try {
                    $originalDefaultPrinter = Change-DefaultPrinter -printerName $printerName
                } catch {
                    $errvariable =  $_
                }
            }

            <#try {
                Start-execPrinter -pdfFile $pdfFile.FullName
            } catch {
                    $errvariable =  $_
            } #>

            if ($printerName) {
                try {
                    Return-DefaultPrinter -originalDefaultPrinter $originalDefaultPrinter[1].Name
                } catch {
                    $errvariable =  $_
                }
            }
            
            if ($SmtpServer) {
                 # Send en mail

                if ($errvariable) {
                    $subject = "MED FEJL! Der er printet en fil - sendt til printer '$printerName'"
                    $body = "Den printede fil er vedhæftet. Der er sket en fejl i processen: '$errvariable'"
                } else{
                    $subject = "Der er printet en fil - sendt til printer '$printerName'"
                    $body = "Den printede fil er vedhæftet."
                }
                try {
                    Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject $subject  -Body $body -Attachments $pdfFile.FullName
                } catch {
                        if (-not $errvariable) {
                            $errvariable = $_
                        }
                        $body = "Mail for printjob med vedhæftning har fejlet pga. '$errvariable'"
                        Send-MailMessage -SmtpServer $SmtpServer -From $senderEmail -To $recipientEmail -Subject "Printjob har fejlet pga. '$errvariable'"  -Body $body
                }
            }
            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force #Rykker til arkiv
        }
    }
}

Export-ModuleMember -Function Start-SendPDFtoPrint



Import-Module -Name C:\Users\madsc\OneDrive\Skrivebord\MchWork\powershellPrintJobs\psModules\SendPDFtoPrint.psm1 -Force

#Tilføj rettigheder til at importere scripts. Udkommenter den ene linje under.
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

###############REDIGER VARIABLE #######################

$smptServer = "din.smtp.server.dk" 

$ModtagerMail = "abc@123.dk"  #Mail, der skal modtage notifikationer

$sourceFolder = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest"

$destinationFolder = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\Arkiverede rapporter"

#$excelPath = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\test.xlsx"  #Excel til oversigt over dine historiske printjobs. Kommenter ud hvis det ikke ønskes

#$printerName = "" #Navn på Printer

############### REDIGER VARIABLE SLUT #######################

# FileSystemWatcher object
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $sourceFolder

# Kun tjek for filer
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName

# Afgræns til pdf og sæt actions
$action = {
    $file = $Event.SourceEventArgs.Name
    if ($file -match '\.pdf') {
        Write-Host "PDF fil fundet: '$pdfFiles'"

        # Sender job til at maile, pakke og dokumentere print
         Start-SendPDFtoPrint -SmtpServer $smptServer  -recipientEmail $ModtagerMail  -sourceFolder $sourceFolder -destinationFolder $destinationFolder 
             #Ikke brugte inputs:
            #$senderEmail - Default: PrintJobs
            #$excelPath - Skal defineres øverst og inkluderes
            #$printerName - Kun hvis der skal bruges en ikke-standard printer.
    }
}

# Registrer event
Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier PDFFileCreated -Action $action

# Start monitoring
$watcher.EnableRaisingEvents = $true

# Scriptet kører hvert x. sekund
try {
    while ($true) {
        # Do nothing to keep the script running
        Start-Sleep -Seconds 120
    }
}
finally {
    # Luk ned når vi er færdige
    Unregister-Event -SourceIdentifier PDFFileCreated
    $watcher.Dispose()
}

    


  

  
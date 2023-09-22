Import-Module -Name C:\Users\madsc\OneDrive\Skrivebord\MchWork\powershellPrintJobs\psModules\SendPDFtoPrint.psm1 -Force

#Tilføj rettigheder til at importere scripts. Udkommenter den ene linje under.
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned


#Ikke den mest sofistikerede løsning, men 

###############REDIGER VARIABLE #######################

$smptServer = "din.smtp.server.dk" 

$ModtagerMail = "abc@123.dk"  #Mail, der skal modtage notifikationer

$sourceFolder = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest"

$destinationFolder = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\Arkiverede rapporter"

$logPath = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\logbog.xlsx"  #Excel til oversigt over dine historiske printjobs. Kommenter ud hvis det ikke ønskes

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
        # Sender job til at maile, pakke og dokumentere print
         Start-SendPDFtoPrint -recipientEmail $ModtagerMail  -sourceFolder $sourceFolder -destinationFolder $destinationFolder -logPath $logPath  #Testvariabel til ikke fysisk at printe eller ligge i excel

            #Ikke brugte inputs:
                #$senderEmail - Default: PrintJobs
                #$printerName - Kun hvis der skal bruges en ikke-standard printer.
                #$smtpserver - Tilføjes hvis man vil sende en mail. 
                #$user - Sættes default til brugeren, der kører scriptet
    }
}

# Registrer event
Register-ObjectEvent -InputObject $watcher -EventName Created -SourceIdentifier PDFFileCreated -Action $action

# Start monitoring
$watcher.EnableRaisingEvents = $true
Write-Host "Venter på nye filer..."
Write-Host "Tryk Ctrl+C for at stoppe processen."
try {
    while ($true) {
        # Sov i x minutter førend det kører igen
        Start-Sleep -Seconds 10
    }
}
finally {
    # Luk ned når vi er færdige
    Unregister-Event -SourceIdentifier PDFFileCreated
    $watcher.Dispose()
}

    


  

  
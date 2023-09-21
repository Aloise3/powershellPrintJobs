Import-Module -Name C:\Users\madsc\OneDrive\Skrivebord\MchWork\powershellPrintJobs\psModules\NotificationFromPrinter.psm1 -Force

#Tilføj rettigheder til at importere scripts. Udkommenter den ene linje under.
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

###############REDIGER VARIABLE #######################

$StartEllerSletProces = 1  # Sættes til 0 hvis du vil fjerne processen

$WMIjobNavn = "NotificationFromPrinter" #Navn på proces der skal laves. Bruges til at lukke den ned igen, hvis der ikke længere er behov for det.

$ModtagerMail = "abc@123.dk"  #Mail, der skal modtage notifikationer

$smptServer = "din.smtp.server.dk" 

$excelPath = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\test.xlsx"  #Excel til oversigt over dine historiske printjobs. Kommenter ud hvis det ikke ønskes

#$tempPath  = "C:/temp"  #Sti til midlertidig copy-paste fra printjob. Bruges ikke i nuværende setup

############### REDIGER VARIABLE SLUT #######################

if ($StartEllerSletProces -eq 1) {

    Start-PrintJobMonitor {
    -SmtpServer $smptServer #SMTP-server. Rediger ovenfor. 
    -userName $Env:UserName  #Brugerens navn
       if ($excelPath) {         #Laver kun et excelark, hvis det udfyldes foroven.
            -ExcelFilePath $excelPath 
            } 
    -SourceIdentifier $WMIjobNavn #Navn på job. Rediger ovenfor.
    -recipientEmail $ModtagerMail #Modtagermail. Rediger ovenfor.
    #Ikke brugte inputs:
    #senderEmail - Default: PrintJobs
    #tempPath - Ikke brugbar. Kræver oversættelse af spoolerkommandoer og det kræver tid og indsigt i det specifikke printersetup. Variabel er en midlertidig sti til opbevaring af fil inden den kopieres ind i mail. Evt. bare sæt sti til excelarkets sti.
    }


}   else {
        Stop-PrintJobMonitor -NameOfProces $WMIjobNavn
    }
Import-Module -Name C:\Users\madsc\OneDrive\Skrivebord\MyModules\psModules\SendPDFtoPrint.psm1 -Force

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


    Start-SendPDFtoPrint -SmtpServer $smptServer  -recipientEmail $ModtagerMail  -sourceFolder $sourceFolder -destinationFolder $destinationFolder -printerName $printerName

    #Ikke brugte inputs:
    #$senderEmail - Default: PrintJobs
    #$excelPath - Skal defineres øverst og inkluderes
    #$printerName - Kun hvis der skal bruges en ikke-standard printer.
  
# powershellPrintJobs

## Generel info
Dette modul viser anvendelsen af to forskellige metoder til at overvåge og opsnappe information om printerejobs og tilhørende filer. 

### Metode 1: WMI-Event til logning og mailing (..\InstanceCreationMonitor\WMI_Job.ps1)

Denne metode opsnapper Events, der sendes til printere, som computeren er forbundet til og opsamler filnavn, tidspunkt og bruger i en logfil (.xlsx). 
Derefter sendes en mail med relevante informationer.

[Hvad er et WMI-event?](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/register-wmievent?view=powershell-5.1)

#### Registering af job

Man åbner ..\InstanceCreationMonitor\WMI_Job.ps1 og tilretter nedenstående sti til den korrekte for modulet NotificationFromPrinter. 
```powershell
Import-Module -Name C:\Users\Skrivebord\powershellPrintJobs\InstanceCreationMonitor\Module\NotificationFromPrinter.psm1 -Force
```
Dette modul indeholder funktionerne ```Start-PrintJobMonitor``` og ```Stop-PrintJobMonitor```

Powershell kalder funktionsargumenter med -'Argument', hvorefter at input skrives ind efter. Variable klassificeres med et '$' foran navnet. 
Mail sendes med powershell-funktionen [Send-MailMessage](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.3).
Syntaksen er således:

```powershell
#Eksempel på input
Start-PrintJobMonitor -userName "TestBruger" -logpath "C:\Users\madsc\OneDrive\Skrivebord\PrintTest" -SourceIdentifier NotificationFromPrinter

#<Forklaring af samtlige inputvariable
    SourceIdentifier: Navn på WMI-jobbet. Dette bruges til at identificere processen, så den kan lukkes ned igen mm.

    userName: Brugeren, som printjobbet skal valideres imod. Den sender kun en mail til brugeren, som har printet filen. Default: Brugernavn på den bruger, der sætter jobbet op.

    logpath: Stien til logfilen, som enten kan være af .xlsx eller .txt format. Stien skal indeholde filnavnet, og filen skal eksistere. 

    SmtpServer: Navnet på SMTP-serveren, som mailen skal sendes gennem. Kræves for at sende en mail. 

    recipientEmail: Navnet på den mail, der skal modtage emailen. Funktionen skal konfigureres til at linke brugernavne op til mails, hvis der skal sendes mails ud til flere brugere.

    senderEmail: Afsender på mail. Behøver ikke at være en valid mailadresse. Default: Printjobs 
#>
```
Da vi nu kender funktionens inputs skal variable sættes. I WMI_Job.ps1 køres koden:

```powershell
#Importerer funktionen til at overvåge printerjobs
Import-Module -Name C:\Users\Skrivebord\powershellPrintJobs\InstanceCreationMonitor\Module\NotificationFromPrinter.psm1 -Force

###############REDIGER VARIABLE #######################

$StartEllerSletProces = 1  # Sættes til 0 hvis du vil fjerne processen. Er bare blevet brugt til tests uden at lukke sessionen...

$WMIjobNavn = "NotificationFromPrinter" #Navn på proces der skal laves. Bruges til at lukke den ned igen, hvis der ikke længere er behov for det.

$ModtagerMail = "abc@123.dk"  #Mail, der skal modtage notifikationer

$smptServer = "din.smtp.server.dk" #SMTP-server. Kræver eventuelt også, at der logges ind. 

$logpath = "C:\Users\madsc\OneDrive\Skrivebord\PrintTest\logbog.xlsx"  #Excel til oversigt over dine historiske printjobs. Kommenter ud hvis det ikke ønskes

############### REDIGER VARIABLE SLUT #######################

#Tjekker om vi starter eller slutter jobbet
if ($StartEllerSletProces -eq 1) {
    #Vi igangsætter WMI-processen
    Start-PrintJobMonitor   -userName $Env:UserName -logpath $logpath -SourceIdentifier $WMIjobNavn -recipientEmail $ModtagerMail -SmtpServer -senderEmail


}   else {
    #Stop WMI-proces. Den stoppes også, når computeren slukker eller powershell-sessionen lukkes.
        Stop-PrintJobMonitor -NameOfProces $WMIjobNavn
    }
```
#### Skal sættes op i task scheduler
Da et WMI-event lukker ned samtidigt med powershell-sessionen kan det være fordelagtigt at opsætte et task-scheduler job, der sættes i gang, når computeren starter, når man logger ind igen og på givne intervaller. 

Se derfor [Opsætning af task i scheduler](#Opsætning-af-task-i-scheduler) for opsætning.

### Metode 2: Opsamling af PDF-print til videre distribuering (..\PDFMonitor\WatcherPDF.ps1)

Denne metode antager, at man initielt laver print til pdf, og ligger pdf-filen i en mappe, der observeres af et program. På et arbitrært interval tjekkes folderen for nye pdf filer. Eksemplet tjekker hvert 10. sekund, men hver 2-5 minutter er nok mere realistisk i virkeligheden.

Findes en (eller flere) ny pdf-fil logges det i et excel-ark, samt sendes et fysisk printerjob til en navngivet printer. 
Til sidst sendes en notifikations-mail med den vedhæftede fil og filen rykkes over i en 'Arkiv'-mappe.

Denne metode fører dermed filen med over i email-notifikationen, men har den begrænsning, at man skal printe en pdf-fil på et bestemt drev. Ellers klares resten automatisk.


OBS: Begge jobs skal sættes op i task scheduler. Dette da, selvom at de er kontinuerte funktioner, så kan sessionen i sjældne tilfælde slukkes. Også der er PDF -> Printer metoden bedre.






## Opsætning af task i scheduler

### Åbn Opgavestyring:

Tryk på Win + S for at åbne Windows-søgefeltet.
Skriv "Opgavestyring" og tryk på Enter.

### Opret en Ny Opgave:

I Opgaveplanlægger-vinduet skal du klikke på "Opret basisopgave..." eller "Opret opgave..." i højre rude. 

![Alt Text](pics\HovedvindueTaskScheduler.png)


### Konfigurér triggers:

I fanen udløsere skal du definere, hvornår og hvor ofte du vil have, at opgaven skal køre. Du kan vælge mellem indstillinger som "Dagligt," "Ugentligt," "Månedligt" eller "Ved logon." Angiv startdato og -tidspunkt.

### Konfigurér Handlinger:

I fanen Handlinger skal du definere, hvad opgaven skal gøre, når den køres. Du kan vælge at starte et program, sende en e-mail, vise en meddelelse og mere. Hvis du vil køre en PowerShell-script, skal du vælge "Start et program" og angive stien til powershell.exe samt scriptet som et argument.

Ofte ligger powershell på "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" og i Argumentet tilføjet scriptet, der skal køres "C:\Path\To\YourScript.ps1".

Eksempel på at køre det usynligt i baggrunden:

-ExecutionPolicy Bypass -File C:\Users\madsc\OneDrive\Skrivebord\MyModules\Scripts\TestPrint

Sættes til at køre ved opstart + en gang per x antal minutter




[def]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/register-wmievent?view=powershell-5.1
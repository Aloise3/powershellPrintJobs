# powershellPrintJobs

** OBS: Lavet i en hurtig vending privat og fungerer på eget lokale miljø. Kræver tilpasning af større eller mindre art afhængig af serveropsætning. Jeg har heller ikke en smtp-server eller en fysisk printer, men de to funktioner er ret ligetil at opstille. Kræver muligvis nogle administraterrettigheder til at ændre printer osv. **

## Generel info
Dette modul viser to forskellige metoder til at overvåge og opsnappe information om printerejobs og tilhørende filer. 

### Metode 1: WMI-Event til logning og mailing (..\InstanceCreationMonitor\WMI_Job.ps1)

Denne metode opsnapper Events, der sendes til printere, som computeren er forbundet til og opsamler filnavn, tidspunkt og bruger i en logfil (.xlsx). 
Derefter sendes en mail med relevante informationer.

Metoden har den begrænsning, at det er et enormt projekt at genskabe selve filen, så derfor kan filen ikke vedhæftes en notifikationsmail. 

Kan derfor kun bruges til dokumentation af, at man har sat et printerjob igang

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
Følg Guiden:

Hvis du valgte "Opret basisopgave," vil en guide guide dig gennem processen. Følg vejledningen og angiv nødvendige oplysninger som opgavens navn og beskrivelse.

Hvis du valgte "Opret opgave," får du mere avancerede indstillinger. Udfyld fanerne Generelt, Udløsere, Handlinger og Betingelser med de passende indstillinger for din opgave.

### Konfigurér triggers:

I fanen udløsere skal du definere, hvornår og hvor ofte du vil have, at opgaven skal køre. Du kan vælge mellem indstillinger som "Dagligt," "Ugentligt," "Månedligt" eller "Ved logon." Angiv startdato og -tidspunkt.

### Konfigurér Handlinger:

I fanen Handlinger skal du definere, hvad opgaven skal gøre, når den køres. Du kan vælge at starte et program, sende en e-mail, vise en meddelelse og mere. Hvis du vil køre en PowerShell-script, skal du vælge "Start et program" og angive stien til powershell.exe samt scriptet som et argument.

Ofte ligger powershell på "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" og i Argumentet tilføjet scriptet, der skal køres "C:\Path\To\YourScript.ps1".

Eksempel på at køre det usynligt i baggrunden:

-ExecutionPolicy Bypass -File C:\Users\madsc\OneDrive\Skrivebord\MyModules\Scripts\TestPrint

Sættes til at køre ved opstart + en gang per x antal minutter



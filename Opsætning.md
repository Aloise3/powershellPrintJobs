##Åbn Opgavestyring:

Tryk på Win + S for at åbne Windows-søgefeltet.
Skriv "Opgavestyring" og tryk på Enter.

##Opret en Ny Opgave:

I Opgaveplanlægger-vinduet skal du klikke på "Opret basisopgave..." eller "Opret opgave..." i højre rude. 
Følg Guiden:

Hvis du valgte "Opret basisopgave," vil en guide guide dig gennem processen. Følg vejledningen og angiv nødvendige oplysninger som opgavens navn og beskrivelse.

Hvis du valgte "Opret opgave," får du mere avancerede indstillinger. Udfyld fanerne Generelt, Udløsere, Handlinger og Betingelser med de passende indstillinger for din opgave.

##Konfigurér triggers:

I fanen udløsere skal du definere, hvornår og hvor ofte du vil have, at opgaven skal køre. Du kan vælge mellem indstillinger som "Dagligt," "Ugentligt," "Månedligt" eller "Ved logon." Angiv startdato og -tidspunkt.

##Konfigurér Handlinger:

I fanen Handlinger skal du definere, hvad opgaven skal gøre, når den køres. Du kan vælge at starte et program, sende en e-mail, vise en meddelelse og mere. Hvis du vil køre en PowerShell-script, skal du vælge "Start et program" og angive stien til powershell.exe samt scriptet som et argument.

Ofte ligger powershell på "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" og i Argumentet tilføjet scriptet, der skal køres "C:\Path\To\YourScript.ps1".

Eksempel på at køre det usynligt i baggrunden:

-ExecutionPolicy Bypass -File C:\Users\madsc\OneDrive\Skrivebord\MyModules\Scripts\TestPrint

Sættes til at køre ved opstart + en gang per x antal minutter


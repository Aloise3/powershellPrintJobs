

#Tilføj rettigheder til at scripte i WMI. Udkommenter den ene linje under. Brug kun hvis det fejler. Åbner PC'en op for scripts, som du godkender.
#Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

# WMI query
$query = @"
SELECT * FROM __InstanceCreationEvent WITHIN 1
WHERE TargetInstance ISA 'Win32_PrintJob'
"@

# Registrer event
Register-WmiEvent -Query $query -SourceIdentifier NewPrintJobEvent -Action {
    $eventArgs = $Event.SourceEventArgs.NewEvent
    $printJob = $eventArgs.TargetInstance

    # Relevante informationerC
    $jobId = $printJob.JobId
    $document = $printJob.Document
    $owner = $printJob.Owner
    $timestamp = Get-Date -Format "yyyy_MM_dd" #tidspunkt

        


    # Output 
    Write-Host "Der printes en ny fil:"
    Write-Host "Job ID: $jobId"
    Write-Host "Dokument: $document"
    Write-Host "Ejer: $owner"
    Write-Host "Tidspunkt: $timestamp"
  
}

# Simpel funktion til at stoppe processen mens man tester
Write-Host "Venter på nye jobs..."
Write-Host "Tryk Ctrl+C for at stoppe processen."
try {
    while ($true) {
        Wait-Event -SourceIdentifier NewPrintJobEvent | Out-Null
    }
} finally {
    Unregister-Event -SourceIdentifier NewPrintJobEvent
    Write-Host "Proces afsluttet."
}

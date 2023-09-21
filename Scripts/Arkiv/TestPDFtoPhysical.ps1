



# Define source and destination folders
$sourceFolder = "C:\SourceFolder"
$destinationFolder = "C:\DestinationFolder"

# Define the name of the physical printer
$printerName = "YourPrinterName"

# Define SMTP server settings (just the server name)
$smtpServer = "smtp.example.com"
$smtpFrom = "your_email@example.com"
$smtpTo = "recipient_email@example.com"
$smtpSubject = "PDF File Confirmation"

# Infinite loop to continuously check for new files
while ($true) {
    # Get a list of PDF files in the source folder
    $pdfFiles = Get-ChildItem -Path $sourceFolder -Filter *.pdf

    # Check if any PDF files are found
    if ($pdfFiles.Count -gt 0) {
        # Loop through each PDF file found
        foreach ($pdfFile in $pdfFiles) {
            # Move the PDF file to the destination folder
            Move-Item -Path $pdfFile.FullName -Destination $destinationFolder -Force
            
            # Print the PDF file to the physical printer
            Start-Process -FilePath $pdfFile.FullName -Verb Print -PassThru | ForEach-Object {
                # Wait for the print job to complete (optional)
                $_ | Wait-PrinterJob
            }
            
            # Send a confirmation email with the PDF file attached
            Send-MailMessage -SmtpServer $smtpServer -From $smtpFrom -To $smtpTo -Subject $smtpSubject -Body "Attached is the PDF file that was printed." -Attachments $pdfFile.FullName
            
            # Optionally, you can delete the file from the source folder after printing and sending the email
            # Remove-Item -Path $pdfFile.FullName -Force
        }
    }

    # Sleep for a specified interval before checking again (e.g., 5 seconds)
    Start-Sleep -Seconds 5
}

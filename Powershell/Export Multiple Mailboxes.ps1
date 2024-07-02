Write-Host "# ------------------------- #
# Kyle Whittle              #
# Optimus Systems           #
# Windows Updates           #
# v1.1                      #
# Created: 12 Dec 2023      #
# Last Updated: 12 Dec 2023 #
# ------------------------- #" -ForegroundColor cyan
# v1.0 - initial script build and testing

# PowerShell script to read email addresses from a CSV file and write them to an export file

# Prompt user for CSV file location
$csvFileLocation = Read-Host "Enter the path to the CSV file (e.g., C:\path\to\input.csv)"

# Prompt user for export location
$exportLocation = Read-Host "Enter the path to export the mailboxes (e.g., C:\path\to\folder)!NO TRAILING BACKSLASH!"

# Prompt user for ticket details
$ticketNumber = Read-Host "Enter the ticket number and a brief summary (e.g, #123456 Archiving Mailboxes)"

# Read the CSV file
$csvData = Import-Csv -Path $csvFileLocation

# Initialize an empty array to store email addresses
$emailAddresses = @()

# Loop through each row in the CSV
foreach ($row in $csvData) {
    # Get the email address from the "email" column
    $email = $row.email
    # Add the email address to the array
    $emailAddresses += $email
}

# Check if the export tool is installed for the user, and download if not.
While (-Not ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)){
    Write-Host "Downloading Unified Export Tool ."
    Write-Host "This is installed per-user by the Click-Once installer."
    # Credit to Jos Verlinde for his code in Load-ExchangeMFA in the Powershell Gallery! All I've done is update the manifest url and remove all the comments
    # Ripped from https://www.powershellgallery.com/packages/Load-ExchangeMFA/1.2
    # In case anyone else has any ClickOnce applications they'd like to automate the install for:
    # If you're looking for where to find a manifest URL, once you have run the ClickOnce application at least once on your computer, the url for the application manifest can be found in the Windows Registry at "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall" (yes, CTR apps are installed per-user).
    # Look through the keys with names that are 16 characters long hex strings. They'll have a string value (REG_SZ) named either "ShortcutAppId" or "UrlUpdateInfo" that contains the URL as the first part of the string.
    $Manifest = "https://complianceclientsdf.blob.core.windows.net/v16/Microsoft.Office.Client.Discovery.UnifiedExportTool.application"
    $ElevatePermissions = $true
    Try {
        Add-Type -AssemblyName System.Deployment
        Write-Host "Starting installation of ClickOnce Application $Manifest "
        $RemoteURI = [URI]::New( $Manifest , [UriKind]::Absolute)
        if (-not  $Manifest){
            throw "Invalid ConnectionUri parameter '$ConnectionUri'"
        }
        $HostingManager = New-Object System.Deployment.Application.InPlaceHostingManager -ArgumentList $RemoteURI , $False
        Register-ObjectEvent -InputObject $HostingManager -EventName GetManifestCompleted -Action { 
            new-event -SourceIdentifier "ManifestDownloadComplete"
        } | Out-Null
        Register-ObjectEvent -InputObject $HostingManager -EventName DownloadApplicationCompleted -Action { 
            new-event -SourceIdentifier "DownloadApplicationCompleted"
        } | Out-Null
        $HostingManager.GetManifestAsync()
        $event = Wait-Event -SourceIdentifier "ManifestDownloadComplete" -Timeout 15
        if ($event ) {
            $event | Remove-Event
            Write-Host "ClickOnce Manifest Download Completed"
            $HostingManager.AssertApplicationRequirements($ElevatePermissions)
            $HostingManager.DownloadApplicationAsync()
            $event = Wait-Event -SourceIdentifier "DownloadApplicationCompleted" -Timeout 60
            if ($event ) {
                $event | Remove-Event
                Write-Host "ClickOnce Application Download Completed"
            }
            else {
                Write-error "ClickOnce Application Download did not complete in time (60s)"
            }
        }
        else {
            Write-error "ClickOnce Manifest Download did not complete in time (15s)"
        }
    }
    finally {
        Get-EventSubscriber|? {$_.SourceObject.ToString() -eq 'System.Deployment.Application.InPlaceHostingManager'} | Unregister-Event
    }
}

# Find the Unified Export Tool's location and create a variable for it
$ExportExe = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter microsoft.office.client.discovery.unifiedexporttool.exe -Recurse).FullName | Where-Object{ $_ -notmatch "_none_" } | Select-Object -First 1)

Write-Host "Connecting to Exchange Online. Enter your admin credentials in the pop-up window."
Connect-IPPSSession

foreach ($emailAddress in $emailAddresses) {
    # Initialize the search status
    $SearchStatus = "NotStarted"
    $ExportName = $emailAddress
    $SearchName = $emailAddress
    
    # Create the compliance search
    New-ComplianceSearch -Name $emailAddress -ExchangeLocation $emailAddress -Description $ticketNumber -AllowNotFoundExchangeLocationsEnabled $true

    # Start the compliance search
    Start-ComplianceSearch -Identity $emailAddress

    # Wait for the compliance search to complete
    while ($SearchStatus -notlike "Completed") {
        Start-Sleep -s 2
        $SearchStatus = Get-ComplianceSearch $emailAddress | Select-Object -ExpandProperty Status
        Write-Host -NoNewline "."
    }

    # Once the search is complete, proceed to create an export from the search
    Write-Host "Compliance search is complete! Creating export from the search..."
    # Add your export code here (e.g., exporting to a PST file)

    # Print a success message
    Write-Host "Search exported - ready for download"
    Write-Host "Creating export from the search..."
    New-ComplianceSearchAction -SearchName $SearchName -Export -Format FxStream -ExchangeArchiveFormat PerUserPst -Scope BothIndexedAndUnindexedItems -Force 
    Write-Host "Pausing script for 2 minutes, waiting for export to be ready to download"
    Start-Sleep -s 120 # Arbitrarily wait 5 seconds to give microsoft's side time to create the SearchAction before the next commands try to run against it. I /COULD/ do a for loop and check, but it's really not worth it.

    $ExportName += "_Export"
    $ExportDetails = Get-ComplianceSearchAction -Identity $ExportName -IncludeCredential -Details
    $ExportDetails = $ExportDetails.Results.split(";")
    $ExportContainerUrl = $ExportDetails[0].trimStart("Container url: ")
    $ExportSasToken = $ExportDetails[1].trimStart(" SAS token: ")
    $ExportEstSize = [double]::Parse($ExportDetails[18].TrimStart(" Total estimated bytes: "))
    $ExportTransferred = [double]::Parse($ExportDetails[20].TrimStart(" Total transferred bytes: "))
    $ExportProgress = $ExportDetails[22].TrimStart(" Progress: ").TrimEnd("%")
    $ExportStatus = $ExportDetails[25].TrimStart(" Export status: ")

    # Download the exported files from Office 365
    Write-Host "Initiating download for $ExportName"
    Write-Host "Saving export to: $exportLocation"
    $Arguments = "-name ""$ExportName""","-source ""$ExportContainerUrl""","-key ""$ExportSasToken""","-dest ""$exportLocation""","-trace true"
    Start-Process -FilePath "$ExportExe" -ArgumentList $Arguments

    # Wait for the process to start
    $started = $false
    do {
        $status = Get-Process microsoft.office.client.discovery.unifiedexporttool -ErrorAction SilentlyContinue
        if (!$status) {
            Write-Host 'Waiting for process to start'
            Start-Sleep -Seconds 5
        } else {
            Write-Host 'Process has started'
            $started = $true
        }
    } until ($started)

    # Monitor export progress
    do {
        $ProcessesFound = Get-Process | Where-Object { $_.Name -like "*unifiedexporttool*" }
        if ($ProcessesFound) {
            $exportDetails = Get-ComplianceSearchAction -Identity $exportname -IncludeCredential -Details
            $Downloaded = $exportDetails.Results.Progress
            $PercentComplete = [math]::Min(($Downloaded / $ExportEstSize * 100), 100)
            Write-Progress -Activity "Export in progress" -Status "Downloading" -PercentComplete $PercentComplete
            Write-Host "Export still downloading, progress is $PercentComplete%, waiting 60 seconds"
            Start-Sleep -Seconds 60
        }
    } until (!$ProcessesFound)
}

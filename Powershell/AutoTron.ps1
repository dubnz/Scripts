Write-Host "# ------------------------- #
# Kyle Whittle              #
# Optimus Systems           #
# Windows Updates           #
# v1.0                      #
# Created: 03 Jul 2024      #
# Last Updated: 03 Jul 2024 #
# ------------------------- #" -ForegroundColor cyan
# v1.0 - Initial build of script

# Bypass execution policy
Set-ExecutionPolicy Bypass -Scope Process -Force

# Check if running as admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
	Write-Host "Please run this script as an administrator."
	pause
	exit
}

#Setting Variables
$directoryPath = "C:\Optimus"
    if (-not (Test-Path -Path $directoryPath)) {
        New-Item -ItemType Directory -Path $directoryPath
        Write-Host "Directory '$directoryPath' created successfully."
    } else {
        Write-Host "Directory '$directoryPath' already exists."
    }
    if (-not (Test-Path -Path $directoryPath\Tron)) {
        New-Item -ItemType Directory -Path $directoryPath\Tron
        Write-Host "Directory '$directoryPath\Tron' created successfully."
    } else {
        Write-Host "Directory '$directoryPath\Tron' already exists."
    }
$file="tron.exe"
Set-Location $directoryPath
echo "Setting location to $directoryPath\Tron" 
Set-Location $directoryPath\Tron
echo "Locating most recent Tron version from BMRF.org Repo"
$url="https://www.bmrf.org/repos/tron/"

#Downloads and Extracts Tron script
$var=(((Invoke-WebRequest -uri $url).links.href | Sort-Object -Descending)[1])
echo "Most recent version in Repo: "
echo $url$var
write-host "Download can take 5-10 minutes" - 
echo "Downloading $var"
Invoke-WebRequest -uri $url$var -OutFile $file
echo "Download Complete" 
"Extracting $file to $directoryPath\Tron"
.\tron.exe
"$file extracted to subfolder $directoryPath\Tron\"
echo "Copying Tron Files to $directoryPath\Tron"
Move-Item -path .\tron\*.bat -Destination $directoryPath\Tron\
Move-item -path .\tron\resources -Destination $directoryPath\Tron\

#Launches Tron in unattended mode and uploads Logs to Vocatus by default.
write-host "Running Tron Script with the following switches 
     - Accept EULA
     - Preserve power settings
     - Skip ALL anti-virus scans
     - Skip defrag
     - Skip de-bloat
     - Skip OneDrive removal"
.\tron.bat -e -p -sa -sd -sdb -sor
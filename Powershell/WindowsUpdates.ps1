Write-Host "# ------------------------- #
# Kyle Whittle              #
# Optimus Systems           #
# Windows Updates           #
# v1.1                      #
# Created: 12 Dec 2023      #
# Last Updated: 12 Dec 2023 #
# ------------------------- #" -ForegroundColor cyan
# v1.0 - Initial build of script for Windows 10 & Windows 11 Updates
# v1.1 - Adding check for update of PSWindowsUpdate and adding Windows Server Updates also

# Bypass execution policy
Set-ExecutionPolicy Bypass -Scope Process -Force

# Check if running as admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
{
	Write-Host "Please run this script as an administrator."
	exit
}

# Check if PSWindowsUpdate module is installed, if not, install it
if (-not (Get-Module -Name PSWindowsUpdate -ListAvailable)) {
    Write-Host "Installing PSWindowsUpdate module..."
    Install-Module -Name PSWindowsUpdate -Force -AllowClobber
} else {
    # If module is installed, check for an update
    Write-Host "Checking for updates to PSWindowsUpdate module..."
    Update-Module -Name PSWindowsUpdate -Force
}

# Import PSWindowsUpdate module
Import-Module PSWindowsUpdate

# Check Windows version
$osName = (Get-ComputerInfo | Select-Object -ExpandProperty OsName).ToLower()

if ($osName -like "*windows 10*")
{
	# Run script for Windows 10
	Write-Host "Running script for Windows 10..." -ForegroundColor Green
	Write-Host "..." -ForegroundColor Green
	
	# Run Get-WindowsUpdate to check for updates
	Write-Host "Checking for Windows updates..." -ForegroundColor Magenta
	Write-Host "..." -ForegroundColor Magenta
	$updates = Get-WindowsUpdate -NotCategory "Drivers" -NotTitle "Windows 11"
	
	# Display available updates
	if ($updates.Count -gt 0)
	{
		Write-Host "Available updates:"
		$updates | Format-Table -AutoSize
		
		# Prompt user to install updates
		Read-Host "Press Enter to install the updates..."
		Write-Host "Installing updates..." -BackgroundColor Green -ForegroundColor White
		Get-WindowsUpdate -NotCategory "Drivers" -NotTitle "Windows 11" -Verbose -Install -AcceptAll -IgnoreReboot
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host
	}
	else
	{
		Write-Host "No updates available." -BackgroundColor Red -ForegroundColor White
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host
	}
}
elseif ($osName -like "*windows 11*")
{
	# Run script for Windows 11
	Write-Host "Running script for Windows 11..." -ForegroundColor Green
	Write-Host "..." -ForegroundColor Green
	
	# Run Get-WindowsUpdate to check for updates
	Write-Host "Checking for Windows updates..." -ForegroundColor Magenta
	Write-Host "..." -ForegroundColor Magenta
	$updates = Get-WindowsUpdate -NotCategory "Drivers"
	
	# Display available updates
	if ($updates.Count -gt 0)
	{
		Write-Host "Available updates:"
		$updates | Format-Table -AutoSize
		
		# Prompt user to install updates
		Read-Host "Press Enter to install the updates..."
		Write-Host "Installing updates..." -BackgroundColor Green -ForegroundColor White
		Get-WindowsUpdate -NotCategory "Drivers" -Verbose -Install -AcceptAll -IgnoreReboot
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host	
	}
	else
	{
		Write-Host "No updates available." -BackgroundColor Red -ForegroundColor White
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host
	}
}
elseif ($osName -like "*windows server*")
{
	# Run script for Windows 11
	Write-Host "Running script for Windows Server..." -ForegroundColor Green
	Write-Host "..." -ForegroundColor Green
	
	# Run Get-WindowsUpdate to check for updates
	Write-Host "Checking for Windows updates..." -ForegroundColor Magenta
	Write-Host "..." -ForegroundColor Magenta
	$updates = Get-WindowsUpdate -NotCategory "Drivers"
	
	# Display available updates
	if ($updates.Count -gt 0)
	{
		Write-Host "Available updates:"
		$updates | Format-Table -AutoSize
		
		# Prompt user to install updates
		Read-Host "Press Enter to install the updates..."
		Write-Host "Installing updates..." -BackgroundColor Green -ForegroundColor White
		Get-WindowsUpdate -NotCategory "Drivers" -Verbose -Install -AcceptAll -IgnoreReboot
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host	
	}
	else
	{
		Write-Host "No updates available." -BackgroundColor Red -ForegroundColor White
		Write-Host "Press Enter to continue..." -ForegroundColor Cyan
		Read-Host
	}
}
else
{
	Write-Host "Unsupported Windows version."
}
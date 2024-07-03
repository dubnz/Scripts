## CHECK REGISTRY PATH EXISTS & SET REGISTRY KEY
# Set variables to indicate value and key to set
$RegistryPath = 'HKLM:\Software\Policies\Microsoft\OneDrive'
$Name         = 'EnableSyncAdminReports'
$Value        = '1'
# Create the key if it does not exist
If (-NOT (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force
  }
# Now set the value
New-ItemProperty -Path $RegistryPath -Name $Name -Value $Value -PropertyType DWORD -Force
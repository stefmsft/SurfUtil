# Set debug level "Continue" or "SilentContinue"
$DebugPReference = "Continue"
# Set Verbose level "Continue" or "SilentContinue"
$VerbosePreference = "Continue"
Import-Module .\Import-SurfaceDrivers.psm1 -force -verbose
Get-module Import-SurfaceDrivers
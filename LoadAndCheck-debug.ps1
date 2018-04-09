# Set debug level to Debug
$DebugPreference = "Continue"
# Set Verbose level to Verbose
$VerbosePreference = "Continue"
Import-Module .\Import-SurfaceDrivers.psm1 -force
Get-module Import-SurfaceDrivers
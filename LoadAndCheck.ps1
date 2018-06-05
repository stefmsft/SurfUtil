$VerbosePreference = "Continue"
$VerbosePreference = "SilentlyContinue"

$MPath = "$PSScriptRoot\Import-SurfaceDrivers.psm1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
Get-module Import-SurfaceDrivers

$MPath = "$PSScriptRoot\Manage-BMR.psm1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
Get-module Manage-BMR
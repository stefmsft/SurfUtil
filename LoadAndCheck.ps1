$MPath = "$PSScriptRoot\Import-SurfaceDrivers.psm1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
Get-module Import-SurfaceDrivers
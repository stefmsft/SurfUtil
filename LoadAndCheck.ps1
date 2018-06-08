$VerbosePreference = "Continue"
$VerbosePreference = "SilentlyContinue"

$MPath = "$PSScriptRoot\SurfUtil.psm1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
get-command -module SurfUtil
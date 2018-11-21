$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"

$MPath = "$PSScriptRoot\SurfUtil.psm1"

write-host "Clean Up Jobs"
get-job | Receive-Job
Remove-job *

write-host "Unblocking files"
Unblock-File $MPath
unblock-File "$PSScriptRoot\*.ps1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
get-command -module SurfUtil
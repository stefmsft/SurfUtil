$VerbosePreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"

$MPath = "$PSScriptRoot\SurfUtil.psm1"

write-host "Clean Up Jobs"
get-job | Receive-Job
Remove-job *

#Check for bad dismounted WIM
$Mounted = get-windowsImage -Mounted
foreach ($Mp in $Mounted) {
    $sp = $Mp.ImagePath.split("\")
    $WimName = $sp[$sp.count -1]
    Write-Host "Dismounting $WimName"
    Dismount-windowsImage -Path $MP.path -Discard | out-null
}

write-host "Unblocking files"
Unblock-File $MPath
unblock-File "$PSScriptRoot\*.ps1"
write-host "Loading Module $MPath"
Import-Module $MPath -force
get-command -module SurfUtil
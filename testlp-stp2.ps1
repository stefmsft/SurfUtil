$VerbosePreference = "Continue"
$DebugPreference = "Continue"

$mnt1 = "c:\scratch\surfutil-vnext\test"
$mnt3 = "c:\scratch\surfutil-vnext\boot"
$distrib = "c:\scratch\surfutil-vnext\iso\en_windows_10_business_edition_version_1803_updated_aug_2018_x64_dvd_5d7e729e"
$inswimpath = "$distrib\sources\install.wim"
$btwimpath = "$distrib\sources\boot.wim"
$inslangpath = "$distrib\sources\lang.ini"

$lg = "fr-fr"
$lgrecabpath1 = "F:\Windows Preinstallation Environment\x64\WinPE_OCs\$lg\lp.cab"
$lgrecabpath2 = "F:\Windows Preinstallation Environment\x64\WinPE_OCs\$lg\WinPE-Setup_$lg.cab"
$lgrecabpath3 = "F:\Windows Preinstallation Environment\x64\WinPE_OCs\$lg\WinPE-Setup-client_$lg.cab"

Set-ItemProperty $btwimpath -name IsReadOnly -value $false
Set-ItemProperty $inswimpath -name IsReadOnly -value $false
Set-ItemProperty $inslangpath -name IsReadOnly -value $false

Mount-WindowsImage -Path $mnt3 -ImagePath $btwimpath -Index 2

Add-WindowsPackage -PackagePath $lgrecabpath1 -Path $mnt3 -Verbose

try {
    Add-WindowsPackage -PackagePath $lgrecabpath2 -Path $mnt3 -Verbose
    Add-WindowsPackage -PackagePath $lgrecabpath3 -Path $mnt3 -Verbose
} catch { continue }

xcopy .\lp\1803\fr-fr "$distrib\sources\fr-fr" /cherkyi

Mount-WindowsImage -Path $mnt1 -ImagePath $inswimpath -Index 5

Dism /image:$mnt1 /get-intl /distribution:$distrib
Dism /image:$mnt1 /gen-langini /distribution:$distrib
Dism /image:$mnt1 /set-allIntl:$lg
Dism /image:$mnt1 /get-intl /distribution:$distrib

Xcopy "$distribution\sources\lang.ini" "$mnt3\sources\lang.ini"

Dismount-WindowsImage -Path $mnt3 -Save -Verbose -CheckIntegrity
Dismount-WindowsImage -Path $mnt1 -discard -Verbose

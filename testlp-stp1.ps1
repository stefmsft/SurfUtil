$VerbosePreference = "Continue"
$DebugPreference = "Continue"

$mnt1 = "c:\scratch\surfutil-vnext\test"
$mnt2 = "c:\scratch\surfutil-vnext\winre"
$rewimpath = "$mnt1\Windows\System32\Recovery\Winre.wim"
$distrib = "c:\scratch\surfutil-vnext\iso\en_windows_10_business_edition_version_1803_updated_aug_2018_x64_dvd_5d7e729e"
$inswimpath = "$distrib\sources\install.wim"
$inslangpath = "$distrib\sources\lang.ini"

$lg = "fr-fr"
$lgcabpath = "F:\x64\langpacks\Microsoft-Windows-Client-Language-Pack_x64_$lg.cab"
$lgrecabpath = "F:\Windows Preinstallation Environment\x64\WinPE_OCs\$lg\lp.cab"

Set-ItemProperty $inswimpath -name IsReadOnly -value $false
Set-ItemProperty $inslangpath -name IsReadOnly -value $false

Mount-WindowsImage -Path $mnt1 -ImagePath $inswimpath -Index 5
Mount-WindowsImage -Path $mnt2 -ImagePath $rewimpath -Index 1

Get-WindowsPackage -Path $mnt1
Add-WindowsPackage -PackagePath $lgcabpath -Path $mnt1 -Verbose
Add-WindowsPackage -PackagePath $lgrecabpath -Path $mnt2 -Verbose

Dism /image:$mnt1 /set-allIntl:$lg
Dism /image:$mnt2 /set-allIntl:$lg

Set-ItemProperty $inslangpath -name IsReadOnly -value $false
Dism /image:$mnt1 /gen-langini /distribution:$distrib
Dism /image:$mnt1 /get-intl /distribution:$distrib
Dism /image:$mnt2 /get-intl

Dismount-WindowsImage -Path $mnt2 -Save -Verbose -CheckIntegrity
Dismount-WindowsImage -Path $mnt1 -Save -Verbose -CheckIntegrity
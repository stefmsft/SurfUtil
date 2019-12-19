    <#
    .DESCRIPTION
    MakeBMR.ps1 allows you the generate an bootable USB Key for a model of Surface. This key will allow you to reinstal the machine in arround 20 min.
    
    .EXAMPLE

    .\MakeBMR.ps1 -Drive D -WindowsVersion 1903 -SurfaceModel "Surface Pro 6" 
    Create a bootable usb key for a Surface Pro 6 with 19H1 (language defaulted to the first ISO file found in the ISO directory)
    .EXAMPLE

    .\MakeBMR.ps1 -Drive D -WindowsVersion 1903 -SurfaceModel "Surface Laptop" -MkISO -InjLP "fr-fr","en-us" -Log -Verbose 
    Create a bootable usb key for a Surface Laptop  with 19H1. I will add language pack for fr and Us. And ISO file will ba also created in the local directory. Verbosity will be on.

    .PARAMETER Drive
    This is the letter of the drive holding your usb key
    .PARAMETER WindowsVersion
    This is the 4 digit version of Windows 10 that you wish to use as a base for the key
    .PARAMETER SurfaceModel
    This is the Surface Model targeted for the key
    .PARAMETER WindowsEdition
    This is the SKU (pro,enterprise, etc ... ) targeted. The default value in Config.xml is "Windows 10 Pro*"
    .PARAMETER MkISO
    Boolean allowing to ask for the creation of a iso file that is a copy of the usb-key. Warning : This lead to a longer creation process.
    .PARAMETER Language
    This is a 2 letter selector (ex: "fr" or "en") allowing the tool to pickup a specific ISO file in the ISO directory. You can have for the same WindowsVersion target 2 or more ISOs available in your directory. The name schema of those file should be LL_windows_10*VVVV*.iso. The language parameter will replace LL in the seek for a valid OS ISO.
    .PARAMETER InjLP
    This is a list of 5 letters string identifying the extra language that will be supported during setup. Language pack injection require some care and preparation steps that are described below.
    .PARAMETER DirectInj
    Boolean that specify if we want to inject the drivers pack directly in the WIM (value $true) or thru a post setup step where the MSI is silently applied (value $False). The later offer a lot of advantage. This is why the default value is $False
    .PARAMETER Log
    oolean triggering a full verbosity of the operations when its value is $True. It is $False by default.
    #>

Param(  [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Drive,
        [int]$WindowsVersion,
        [string]$SurfaceModel,
        [string]$PathToISO,
        [String]$DrvRepo,
        [string]$WindowsEdition,
        [switch]$MkISO,
        [string]$Language,
        [string[]]$InjLP,
        [switch]$DirectInj,
        [switch]$Log,
        [switch]$Yes
    )

Import-Module "$PSScriptRoot\SurfUtil.psm1" -force | Out-Null

try {

    #Verifiy if ran in Admin
    $IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    if ($IsAdmin -eq $False) {
        Write-Host -ForegroundColor Red "Please use this script in an elevated Admin context"
        return $false
    }


    if ($Log -eq $true) {

        $OldVerboseLevel = $VerbosePreference
        $OldDebugLevel = $DebugPreference

        $VerbosePreference = "Continue"
        $DebugPreference = "Continue"

    }


    $DefaultFromConfigFile = Import-Config
    ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB

    if ($PathToISO -eq "") {

        $localp = (Get-Item -Path ".\" -Verbose).FullName

        $IsoPath = $DefaultFromConfigFile["RootISO"]
        if ($null -eq $IsoPath) {
            $IsoPath = "$localp\Iso"
        }
        $PathToISO = $IsoPath
    }
    If(!(test-path $PathToISO)) {
        New-Item -ItemType Directory -Force -Path $PathToISO | out-null
    }

    if ($DrvRepo -eq "") {

        $localp = (Get-Item -Path ".\" -Verbose).FullName

        $RepoPath = $DefaultFromConfigFile["RootRepo"]
        if ($null -eq $RepoPath) {
            $RepoPath = "$localp\Repo"
        }
        $DrvRepo = $RepoPath
    }
    If(!(test-path $DrvRepo)) {
        New-Item -ItemType Directory -Force -Path $DrvRepo | out-null
    }

    if ($WindowsVersion -eq "") {

        $ios = (Get-WmiObject Win32_OperatingSystem | select-object BuildNumber).BuildNumber
        $WindowsVersion = (($OSReleaseHT.GetEnumerator()) | Where-Object { $_.Value -eq $ios }).Name

    }

    if ($SurfaceModel -eq "") {

        Write-Verbose "No Surface Model parameter provided ... Looking to the default config file"
        $SurfaceModel = $DefaultFromConfigFile["SurfaceModel"]
        if ($SurfaceModel -eq "") {
            Write-Host -ForegroundColor Red "Please specifiy a Surface Model"
            return $false
        }
    }

    if ($WindowsEdition -eq "") {

        Write-Verbose "No Windows SKU parameter provided ... Looking to the default config file"
        $TargetSKU = $DefaultFromConfigFile["TargetSKU"]
        if ($TargetSKU -eq "") {
            Write-Host -ForegroundColor Red "Please specifiy a Windows Sku"
            return $false
        }
    }

    if ($Language -eq "") {

        Write-Verbose "No Language parameter provided ... Looking to the default config file"
        $Language = $DefaultFromConfigFile["Language"]
        if ($Language -eq "") {
            Write-Verbose "Revert to default - Language=en"
            $Language = "en"
        }
    }

    #Check on the Drive
    $TargetDrv = Get-WmiObject win32_volume|where-object {$_.driveletter -match "$Drive"}
    $TargetSize = [int32]($TargetDrv.Capacity / 1GB)
    $TargetLabel = $TargetDrv.label
    write-verbose "Target Drive size is $TargetSize GB"
    write-verbose "Target Drive label is $TargetLabel"

    if (($TargetSize -lt 4) -Or ($TargetSize -gt 50)) {
        Write-Host -ForegroundColor Red "Use a USB Drive with capacity between 4 and 50 Go";
        return $False
    }

    Write-Host "Create a BMR Key for [$SurfaceModel] / [$TargetSKU $WindowsVersion]"


    Write-Verbose "Calling New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU -Language $Language -InjectLP $InjLP -Log $Log -DirectInj $DirectInj -AutoAccept $Yes"
    $ret = New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU -Language $Language -InjectLP $InjLP -Log $Log -DirectInj $DirectInj -AutoAccept $Yes

    if ($DirectInj -eq $true) {
        if ($ret -eq $true) {
            if (($SurfaceModel.tolower() -eq "surface pro") -or ($SurfaceModel.tolower() -eq "surface pro lte")) {
                write-host "*******************************************************************"
                write-host "* Warning information (Surface Pro/LTE only):                                           *"
                write-host "* Verify that the version of the Surface System Aggregator        *"
                write-host "* firmware of the targeted surface is in version v234.2110.1.0    *"
                write-host "* or greater before using this key                                *"
                write-host "*******************************************************************"
            }
        }
    }

}
catch [System.Exception] {
    Write-Host -ForegroundColor Red $_.Exception.Message;
    return $False
}
finally {
    if ($Log -eq $true) {

        write-verbose "Re establish initial verbosity"
        $VerbosePreference = $OldVerboseLevel
        $DebugPreference = $OldDebugLevel

    }
}
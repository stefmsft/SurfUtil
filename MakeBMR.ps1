Param(  [Parameter( Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Drive,
        [int]$WindowsVersion,
        [string]$SurfaceModel,
        [string]$PathToISO,
        [String]$DrvRepo,
        [string]$WindowsEdition,
        [bool]$MkISO,
        [bool]$Log
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


    Write-Verbose "Calling New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU -Log $Log"
    New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU -Log $Log

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
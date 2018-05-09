Param(  [Parameter( Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Drive,
        [int]$WindowsVersion,
        [string]$SurfaceModel,
        [string]$PathToISO,
        [String]$DrvRepo,
        [string]$WindowsEdition,
        [bool]$MkISO
    )

Import-Module "$PSScriptRoot\Manage-BMR.psm1" -force | Out-Null

try {

    $DefaultFromConfigFile = Import-Config
    ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB

    if ($PathToISO -eq "") {

        $localp = (Get-Item -Path ".\" -Verbose).FullName

        $IsoPath = $DefaultFromConfigFile["RootISO"]
        if ($IsoPath -eq $null) {
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
        if ($RepoPath -eq $null) {
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


    Write-Verbose "Calling New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU"
    New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo -MkIso $MkISO -TargetSKU $TargetSKU

}
catch [System.Exception] {
    Write-Host -ForegroundColor Red $_.Exception.Message;
    return $False
}
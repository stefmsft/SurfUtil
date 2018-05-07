Param([string]$WindowsEdition, [int]$WindowsVersion, [string]$SurfaceModel,[string]$PathToISO,[String]$Drive,[String]$DrvRepo)

Import-Module "$PSScriptRoot\Manage-BMR.psm1" -force

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

        Write-Host "No Surface Model parameter provided ... Looking to the default config file"
        $SurfaceModel = $DefaultFromConfigFile["SurfaceModel"]
        if ($SurfaceModel -eq "") {
            Write-Host -ForegroundColor Red "Please specifiy a Surface Model"
            return $false
        }
    }

    Write-Host "Create a BMR Key for [$SurfaceModel] / [Windows 10 $WindowsVersion]"

    if ($Drive -eq "") {

        Write-Host -ForegroundColor Red "Provide the -Drive parameter as a target for the key to generate"
        return $false

    }

    Write-Verbose "Calling New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo"
    New-USBKey -Drive $Drive -ISOPath $IsoPath -Model $SurfaceModel -OSV $WindowsVersion -DrvRepoPath $DrvRepo

}
catch [System.Exception] {
    Write-Host -ForegroundColor Red $_.Exception.Message;
    return $False
}
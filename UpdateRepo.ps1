Import-Module .\Import-SurfaceDrivers.psm1 -force

$DefaultFromConfigFile = Import-Config
    
$localp = (Get-Item -Path ".\" -Verbose).FullName

$RepoPath = $DefaultFromConfigFile["RootRepo"]
if ($RepoPath -eq $null) {
    $RepoPath = "$localp\Repo"
}

If(!(test-path $RepoPath)) {
        New-Item -ItemType Directory -Force -Path $repopath | out-null
}

$LocalRepoPathDir = resolve-path $RepoPath
write-debug "Full Path Repo is : $LocalRepoPathDir"

try {

    Write-Host $VerbosePreference
    ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB

    foreach ($SurfaceModel in $SurfModelHT.keys) {
        Import-SurfaceDrivers -Model $SurfaceModel -CheckOnly $False -Root $LocalRepoPathDir
    }

}
catch [System.Exception] {
    Write-Host -ForegroundColor Red $_.Exception.Message;
    return $False
}
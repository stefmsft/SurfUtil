function Get-LocalDriversInfo {
    [CmdletBinding()]
    param
    (
        [Alias('RootRepo')]
        [string]$RPath
        ,
        [Alias('Model')]
        [string]$DrvModel
        ,
        [Alias('OSTarget')]
        [string]$OVers
    )
    
    begin {
         Write-Verbose "Begin procesing Get-LocalDriversInfo(Repo=$RPath,Model=$DrvModel,OSVersion=$OVers)"  
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
   }
  
    process {

        try {

            # Verify if DrvInfo is empty ... Would mean no internet connexion during the online check
            if ($Global:CacheMode -eq $true) {

                    write-verbose "Cache mode processing for local drivers info"
                    $intver=0
                    #Search the highest Version number
                    foreach ($Release in $OSReleaseHT.keys) {
                        $curintver = [convert]::ToInt32($Release, 10)
                        if($curintver -gt $intver) {
                            $tp = "$RPAth\$DrvModel\$curintver"
                            # verify if a real path actualy exist for this current higher version
                            If(test-path $tp) {
                                $intver = $curintver        
                            }

                        }
                    }
                    if ($intver -ne 0 ) {
                        write-verbose "Found $intver as the latest version available"
                        $OVers = $intver
                    }

                $lpathtocheck = "$RPAth\$DrvModel\$Overs"
                $lpathtocheck = (Get-Item -Path $lpathtocheck).FullName


                $Dir = get-childitem $lpathtocheck
                $ListExistMSI = $Dir | where-object {$_.extension -eq ".msi"}
                if ($ListExistMSI.count -ne 0) {
                    foreach ($msifile in $ListExistMSI) {
                        $FoundDrvHT = @{} 
                        $ret = $FoundDrvHT.Add("OSVersion",$Overs)
                        $ret = $FoundDrvHT.Add("LPath",$lpathtocheck)
                        $ret = $FoundDrvHT.Add("FileName",$msifile)
                        $ret = $FoundDrvHT.Add('LatestPresent',"Y")
                        $ret = $Global:DrvInfo.Add($FoundDrvHT)
                        write-debug "Inject a local ref inside of DrvInfo"
                        break
                    }
                } else {

                    Write-Host -ForegroundColor Red "No msi found localy and remote driver site is unreachable"
                    return $False
        
                }


            } else {
                foreach ($DrvHashT in $Global:DrvInfo) {

                    $OSSub = $DrvHashT['OSVersion']
                    $LocalPathDir = "$RPath\$DrvModel\$OSSub"
                    # If Overs is specified, we narrow the check to only this version
                    if ($OVers -ne "") {

                        if ($OSSub -ne $OVers) { Break }
                        
                    }
                    write-verbose "Testing if $LocalPathDir exist for [$DrvModel][$OSSub]"
                    If(!(test-path $LocalPathDir)) {

                        write-verbose "Create $LocalPathDir directory"
                        New-Item -ItemType Directory -Force -Path $LocalPathDir | out-null

                    }
                    
                    $FileName = $DrvHashT['FileName']
                    write-verbose "Check for $FileName"
                    $LocalPathFile = "$LocalPathDir\$FileName"
                    $DrvHashT.Add('LPath',$LocalPathDir)
    
                    If(!(test-path $LocalPathFile)) {

                        $DrvHashT.Add('LatestPresent',"N")
                        write-host "$FileName is missing"

                    } else {

                        $DrvHashT.Add('LatestPresent',"Y")
                        write-Verbose "$FileName is already downloaded"                        
                        write-host "Driver for $OSSub found localy"                        

                    }
                }
            }

        return $true
        }
        catch [System.Exception] {

            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False

        }

        return $True
    }

    end {
         Write-Verbose "End procesing Get-LocalDriversInfo"  
    }
}
function Get-RemoteDriversInfo {
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0,mandatory=$true)]
        [Alias('Model')]
        [string]$DrvModel
        ,
        [Alias('OSTarget')]
        [string]$OVers
    )
    
    begin {
        Write-Verbose "Begin procesing Get-RemoteDriversInfo(Model=$DrvModel,OSVersion=$Overs)"  
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
    }
  
    process {

        try {

            [System.Collections.ArrayList]$CurLst = @()
            
            if ($DrvModel -ne "") {
                $urldrv = $SurfModelHT[$DrvModel.tolower()]
                if ($urldrv -eq $null) {
                    Write-Host -ForegroundColor Red "Unknown Surface Model for the script : [$DrvModel]"
                    return $false
                }
            }

            if (($OVers -ne $null) -and ($OVers -ne "")) {
                $InternalR = $OSReleaseHT[$OVers.tolower()]
                if ($InternalR -eq $null) {
                    Write-Host -ForegroundColor Red "Unknown OS Release for the script : [$OVers]"
                    return $false
                }
            }

            Write-Verbose "Processing $urldrv"  

            try {

                $DrvPage = Invoke-WebRequest -Uri $urldrv -UseBasicParsing

            }
            catch [System.Exception] {
                Write-Host "Internet Drivers Site unreachable : Cache mode enabled"
                $Global:CacheMode = $true
                write-verbose "Set Cache Mode to $Global:CacheMode"

                return $false
            }
        
            foreach ($link in $DrvPage.Links) {

                $href = $link.href
                if ($href -ne $null) {
                    if ($href -like "*win10*.msi" ) {
                        $DrvPrsLst = $href.tolower() -split '/'
                        $FileName = $DrvPrsLst[$DrvPrsLst.count-1] -split '.msi'
                        $DriverInfo = $FileName[0] -split '_'
                        $InternalVFound = $DriverInfo[$DriverInfo.count-3]
                        write-debug "Filename identified : $FileName"

                        if ($OSReleaseHT.containsValue($InternalVFound)) {
                            $VFound = $OSReleaseHT.Keys | ForEach-Object { if ($OSReleaseHT.Item($_) -eq $InternalVFound ) {$_} }
                        }
                        else {
                            $VFound = "1507"
                        }

                        if ($VFound -notin $CurLst) {

                            Write-Verbose "[$ret]:$href"   
                            Write-Verbose "Found OS Version is $VFound"
                            if ($VFound -NotIn $CurLst) {
                                if ($OVers -ne "") {
                                    Write-Verbose "Asked OS Version is $OVers"  
                                    if ($Overs -eq $VFound) {
                                        $FoundDrvHT = @{} 
                                        $ret = $FoundDrvHT.Add("OSVersion",$VFound)
                                        $ret = $FoundDrvHT.Add("Link",$href.tolower())
                                        $ret = $FoundDrvHT.Add("FileName",$DrvPrsLst[$DrvPrsLst.count-1])
                                        $ret = $Global:DrvInfo.Add($FoundDrvHT)
                                        write-debug "Remote info for $VFound Added"
                                    }
                                }
                                else {
                                    $FoundDrvHT = @{} 
                                    $ret = $FoundDrvHT.Add("OSVersion",$VFound)
                                    $ret = $FoundDrvHT.Add("Link",$href.tolower())
                                    $ret = $FoundDrvHT.Add("FileName",$DrvPrsLst[$DrvPrsLst.count-1])
                                    $ret = $Global:DrvInfo.Add($FoundDrvHT)
                                    write-debug "Remote info for $VFound Added"
                            }
                            $ret = $CurLst.Add($VFound)
                        }
                    }
                    else { write-verbose "Url skipped : Already in our list"}    
                }
            }
        }
    }
    catch [System.Exception] {
        Write-Host -ForegroundColor Red $_;
        return $false
    }

    return $true
    }

    end {
         Write-Verbose "End procesing Get-RemoteDriversInfo"  
    }
}
function Import-SurfaceDB {
    <#
    .SYNOPSIS
    Helper function exported from the Import-SurfaceDrivers.psm1 module
    This function is made available outside of the module because it can be usefull for other script or modules

    .DESCRIPTION
    This function gather informations present in the ModelsDB.xml and return 2 hash table containing the values
    - The first hash table contains the URL where to download the driver set for each model
    - The second one contains the Internal release numbers for each supported Windows 10 version

    Be carefull to keep the ModelsDB.XML file in the same directory than the module

    The structure of the XML is : 

    <ModelsDB>
        <SurfacesModels>
            <Surface ID='Surface name'>
                <Drivers type='msi' url='url of the download msi page'>
                </Drivers>
                <OS MinSupported='CodeName'>
                </OS>
            </Surface>
        </SurfacesModels>
        <OSRelease>
            <CODENAME ReleaseCode='value' internalCode='value'></CODENAME>
        </OSRelease>
    </ModelsDB>

    .EXAMPLE
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB

    #>
        $DBFileName = "$PSScriptRoot\ModelsDB.xml"
        If(test-path $DBFileName) {
            [XML]$ModelDBFile = Get-Content $DBFileName
        }
        else {
             Write-Verbose "Warning Model DB File not found"  

        }
    
        $ModelsHT = @{}   # empty models hashtable   

        foreach ($Child in $ModelDBFile.ModelsDB.SurfacesModels.ChildNodes ) {

                $ModelsHT.Add($Child.ID.tolower(),$Child.Drivers.url.tolower())
        }

        $OSHT = @{}   # empty models hashtable

        foreach ($Child in $ModelDBFile.ModelsDB.OSRelease.ChildNodes ) {

                $OSHT.Add($Child.ReleaseCode.tolower(),$Child.InternalCode.tolower())
        }

    $CtxData = ($ModelsHT,$OSHT)

    return $CtxData
}
function Import-Config
    {
    <#
    .SYNOPSIS
    Helper function exported from the Import-SurfaceDrivers.psm1 module
    This function is made available outside of the module because it can be usefull for other script or modules

    .DESCRIPTION
    This function gather informations present in the Config.xml and return a hash table containing the values
    The purpose of this config file is mainly to hold the default values for same function parameters

    Be carefull to keep the Config.XML file in the same directory than the module

    The structure of the XML is : 
    
    <Config>
        <Defaults>
            <ParameterName>value</ParameterName>
        </Defaults>
    </Config>

    .EXAMPLE
        $DefaultFromConfigFile = Import-Config

    #>

    $ConfFileName = "$PSScriptRoot\Config.xml"
    If(test-path $ConfFileName) {
        [XML]$ConfigFile = Get-Content $ConfFileName
    }
    else {
         Write-Verbose "Warning Config File not found"  
        return $null
    }

    $Defaults = @{}   # empty models hashtable

    foreach ($Child in $ConfigFile.Config.Defaults.ChildNodes ) {
        $Defaults.Add($Child.Name.tolower(),$Child.InnerText.tolower())
    }

    return $Defaults
}
function Set-DriverRepo {
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0,mandatory=$true)]
        [Alias('Root')]
        [string]$RootRepo
        ,
        [Parameter(mandatory=$true)]
        [Alias('Model')]
        [string[]]$SubFolder
    )
  
    begin {
        Write-Verbose "Begin procesing Set-DriverRepo(Root=$RootRepo,Model=$SubFolder)"  
    }
  
    process {
  
        Write-Verbose "Verify $RootRepo"
        try {

            If(!(test-path $RootRepo))
            {
                Write-Verbose "Create $RootRepo"
                New-Item -ItemType Directory -Force -Path $RootRepo | out-null
            }

            foreach ($s in $SubFolder) {
                Write-Verbose "Processing $s"
                $TstSub = "$RootRepo\$s"
                Write-Verbose "Verify Directory $TstSub"
                If(!(test-path $TstSub))
                {
                    New-Item -ItemType Directory -Force -Path $TstSub | out-null
                }
            }
        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

        return $True

      }

      end {
         Write-Verbose "End procesing Driver Repo"  
    }
}
function Install-MSI{  

    begin {
         Write-Verbose "Begin procesing Install-MSI"  
    }
  
    process {

        try {

            foreach ($DrvHashT in $Global:DrvInfo) {

                $SMsiPath = $DrvHashT['LPath']
                $SMsiFile = $DrvHashT['FileName']
                $SMsiFull = "$SMsiPath\$SMsiFile"
                $TMsiFile = "$env:TEMP\SDrivers.msi"
                $TLogFile = "$env:TEMP\SDrivers.log"

                Copy-Item $SMsiFull -Destination $TMsiFile -Force

                $Arguments = @()
                $Arguments += "/i"
                $Arguments += "$TMsiFile"
                $Arguments += "/qn /norestart"
                $Arguments += "/log $TLogFile"
                
                Write-Host "Applying MSI $MsiFile to the Machine"
                Write-Verbose "Applying MSI Command : MSIEXEC $Arguments"
                Start-Process "msiexec.exe" -ArgumentList $Arguments -Wait
                Write-Host "done"
                Write-Host "Reboot have been blocked, so you might want to reboot your Surface yourself"
            }

        }
        catch [System.Exception] {

            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
         Write-Verbose "End procesing Install-MSI"  
    }
}
function Expand-MSI{  

    begin {
         Write-Verbose "Begin procesing Expand-MSI"  
    }
  
    process {

        try {

            foreach ($DrvHashT in $Global:DrvInfo) {

                $SMsiPath = $DrvHashT['LPath']
                $SMsiFile = $DrvHashT['FileName']
                $SMsiFull = "$SMsiPath\$SMsiFile"
                $TMsiFile = "$env:TEMP\SDrivers.msi"
                $TLogFile = "$env:TEMP\SDrivers.log"

                Copy-Item $SMsiFull -Destination $TMsiFile -Force

                $TmpEMSI = "$env:TEMP\ExpandedMSIDir"

                # Verify Temp Expanded MSI Directory
                write-verbose "Test if $TmpEMSI exist"

                If(!(test-path $TmpEMSI)) {
                    write-verbose "Create $TmpEMSI directory"
                    New-Item -ItemType Directory -Force -Path $TmpEMSI | out-null
                } else {
                    write-verbose "Removing $TmpEMSI content"
                    $FileContent2Delete = "$TmpEMSI\*.*"
                    Remove-Item -Path $FileContent2Delete -Recurse
                }
            
                $Arguments = @()
                $Arguments += "/a"
                $Arguments += "$TMsiFile"
                $Arguments += "targetdir=$TmpEMSI"
                $Arguments += "/qn"
                $Arguments += "/log $TLogFile"
                
                Write-Host "Expanding MSI $MsiFile ...."
                Write-Verbose "Applying MSI Command : MSIEXEC $Arguments"
                Start-Process "msiexec.exe" -ArgumentList $Arguments -Wait
                Write-Host "done"
            }

        }
        catch [System.Exception] {

            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
         Write-Verbose "End procesing Expand-MSI"  
    }
}
function Get-MSIFile {
    <#
    .SYNOPSIS
    TODO
    .DESCRIPTION
    TODO
    .EXAMPLE
    TODO
    .EXAMPLE
    TODO
    .PARAMETER x
    TODO
    #>
    [CmdletBinding()]
    param
    (
        [Alias('Link')]
        [string]$url
        ,
        [Alias('LPath')]
        [string]$targetPath
        ,
        [Alias('File')]
        [string]$targetFile
    )
    
    begin {
         Write-Verbose "Begin procesing Get-MSIFile"  
         Import-Module BitsTransfer
    }
  
    process {

        try {

            $FullTP = resolve-path $TargetPath
            $FulltargetFile = "$FullTP\$TargetFile"
            write-host $FulltargetFile

            Start-BitsTransfer -Source $url -Destination $FulltargetFile

        }
        catch [System.Exception] {

            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
         Write-Verbose "End procesing Get-MSIFile"  
    }
}

function Get-SurfaceDriver {
    [CmdletBinding()]
    param
    (
        [Alias('Apply')]
        [Boolean]$ApplyDRv = $false
    )
    
    begin {
         Write-Verbose "Begin procesing Get-SurfaceDriver(Apply=$ApplyDrv)"  
    }
  
    process {

        try {

            foreach ($DrvHashT in $Global:DrvInfo) {
                $FileName = $DrvHashT['FileName']
                write-verbose "Gathering $FileName ..."
                if ($DrvHashT["LatestPresent"] -eq "N") {
                    $Link = $DrvHashT['Link']
                    $LPath = $DrvHashT['LPath']
                    Write-Verbose "Will load $FileName"
                    Write-Verbose "From $Link"
                    Write-Verbose "To $LPath"

                    $Dir = get-childitem $LPath
                    $ListExistMSI = $Dir | where-object {$_.extension -eq ".msi"}
                    if ($ListExistMSI.count -ne 0) {
                        foreach ($msifile in $ListExistMSI) {
                            $MsiToRemove = "$LPath\$msifile"
                            Remove-Item $MsiToRemove
                        }
                    }

                    #download the Msi file
                    $Strt = Get-Date
                    Write-Host "Start Downloading .................... $FileName"
                    Get-MSIFile -Link $Link -LPath $LPath -File $FileName
                    $End = Get-Date
                    $Span = New-TimeSpan -Start $Strt -End $End
                    $Min = $Span.Minutes
                    $Sec = $Span.Seconds

                    Write-Host "Downloaded in $Min Min and $Sec Seconds"
                } else {

                    write-verbose "File already available locally ..."

                }
            }

        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

        return
    }

    end {
         Write-Verbose "End procesing Get-SurfaceDriver"  
    }
}
function Import-SurfaceDrivers {
    <#
    .SYNOPSIS
    This is the main function exported by the ImportSurfaceDrivers.psm1 module.

    .DESCRIPTION
    This module function allow the import of Surface Driver Set (MSI) from the official external download web site and gather them in an organized way (the repo) so it can be used later by other tools.
    The function held some parameters allowing to apply or expand the driver set on the current machine.
    
    .EXAMPLE
    Import-SurfaceDrivers

    None of the parameters are required, so when it is called like this, the Model Info and the RepoPath parameter will be fetch from the Config.xml file. If this file doesn't exists then it will fail with an explicit error.
    Note that when -Apply or -Expand is not present and no WindowsVersion is targeted, all the available drivers set are downloaded in the local repository.

    .EXAMPLE
    Import-SurfaceDrivers -Model "Surface Pro3" -Expand $True

    The function will fetch the Surface Pro3 drivers Set targeted for the latest Windows 10 version available online.
    Then it will expand the content of the MSI in the directory $ENV:TMP\ExpandedMSIDir

    .EXAMPLE
    Import-SurfaceDrivers -Model "Surface Laptop" -Windowsversion 1709 -Apply $true

    The function will fetch online the Surface Laptop drivers Set targeted for Windows 10 1709.
    If the latest version is already present locally in the repo it won't donwload it.
    Be carrefull, if no version specificaly target the asked version, the function will return an error.
    Then it will expand the content of the MSI in the directory $ENV:TMP\ExpandedMSIDir

    .PARAMETER Model
    Surface Model targeted for the drivers
    
    Supported Models are :
    - Surface Pro
    - Surface Pro Lte
    - Surface Laptop
    - Surface Book 2
    - Surface Book
    - Surface Studio
    - Surface Pro 4
    - Surface Pro 3

    .PARAMETER WindowsVersion    
    Targeted Version of Windows 10
    Supported value are listed in the ModelsDB.XML file.

    .PARAMETER  RepoPath
    Path of the Repo of drivers by models and by Windows version.
    See below an exemple of the structure :
    .\Repo
        Surface model
            Targeted OS Version
                MSI File
    
    .PARAMETER  Apply
    Ask to Apply the driver Set on the actual machine

    .PARAMETER  Expand
    Ask to Only expand the driver Set (Function used by the New-USB function the Manage-BMR module for instance)

    .PARAMETER  CheckOnly
    Only verify if a new download of a driver set is needed but do nothing

    #>
    [CmdletBinding()]
    param
    (
        [string]$Model
        ,
        [string]$WindowsVersion
        ,
        [string]$RepoPath
        ,
        [Boolean]$Apply = $False
        ,
        [Boolean]$Expand = $False
        ,
        [Boolean]$CheckOnly = $False  
    )

    begin {
        Write-Verbose "Begin procesing Import-SurfaceDrivers"
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
        $DefaultFromConfigFile = Import-Config
        $status = $false
    }
  
    process {

        # Verifiy the Surface model input
        if ($Model -eq "") {
            if ('SurfaceModel' -in $DefaultFromConfigFile.Keys) {
                Write-Verbose "Getting Surface Model from Defaults in Config file"  
                $Model = $DefaultFromConfigFile['SurfaceModel']
            }
            else {
                Write-Host -ForegroundColor Red "Surface Model need to be specified in Input or in the Config file"
                return $false
            }
        }

        if ($Model.ToLower() -NotIn $SurfModelHT.keys) {
                Write-Host -ForegroundColor Red "Surface Model $Model not supported by this tool"
            return $false
        }

        # Verifiy the Surface model input
        if ($RepoPath -eq "") {
            if ('RootRepo' -in $DefaultFromConfigFile.Keys) {
                Write-Verbose "Getting Repo path from Defaults in Config file"  
                $RepoPath = $DefaultFromConfigFile['RootRepo']
            }
            else {
                Write-Host -ForegroundColor Red "Repo Path need to be specified in Input or in the Config file"
                return $false
            }
        }
        $RepoPath = (Get-Item -Path $RepoPath -Verbose).FullName
        If(!(test-path $RepoPath)) {

            write-verbose "Create $RepoPath directory"
            New-Item -ItemType Directory -Force -Path $RepoPath | out-null

        }



        # Display the target OS revision asked
        if ($WindowsVersion -ne "") {
            Write-Host "Check Drivers Repo for $Model for Windows $WindowsVersion"
        }
        else {
            Write-Host "Check Drivers Repo for $Model"            
        }

        # Check and prepare the Repo for Model/OS Targeted
        If (Set-DriverRepo -RootRepo $RepoPath -Model $Model) {
             Write-Verbose "Drivers Repo Checked and Set"  
        }
        else {
             Write-Verbose "Error while checking Drivers Repo"
             return $false
        }

        [System.Collections.ArrayList]$Global:DrvInfo = @()
        $Global:CacheMode = $False
        write-verbose "Initializing Cache Mode to $Global:CacheMode"

        # Gather Online drivers Pack reference (Details + urls in a hash table [$Global:DrvInfo])
        # Then if any exist (should always be true), we gather the local picture of the Repo
        if ($WindowsVersion -ne "") {
            $status = Get-RemoteDriversInfo -DrvModel $Model -OSTarget $WindowsVersion
            $status = Get-LocalDriversInfo -RootRepo $RepoPath -DrvModel $Model -OSTarget $WindowsVersion
            if ($Global:DrvInfo.Count -eq 0) {
                Write-Host ">>>   Drivers not found for $WindowsVersion"
                return $false
            }
        }
        else {
            $status = Get-RemoteDriversInfo -DrvModel $Model
            $status = Get-LocalDriversInfo -RootRepo $RepoPath -DrvModel $Model
            if ($Global:DrvInfo.Count -eq 0) {
                Write-Host -ForegroundColor Red "   No Drivers found "
                return $false
            }
        }

        # If Apply is asked we only need the latest driver version
        if ((($Apply -eq $True) -or ($Expand -eq $True)) -and ($Global:DrvInfo.count -gt 1)) {
            #Trim $Global:DrvInfo to the latest Driver only
            write-verbose "Trim on multiple drivers Found"
            $intver=0
            #Search the highest Version number
            foreach ($DrvHashT in $Global:DrvInfo) {
                $curintver = [convert]::ToInt32($DrvHashT['OSVersion'], 10)
                if($curintver -gt $intver) {
                    $HT = $DrvHashT
                    $intver = $curintver
                }
            }
            if ($intver -ne 0 ) {
                write-verbose "Found $intver as the latest version available"
                #reset the Driver list to one entry
                [System.Collections.ArrayList]$Global:DrvInfo = @()
                $Global:DrvInfo.Add($HT) | Out-Null
            }
        }

        if ($CheckOnly -ne $True) {

            Get-SurfaceDriver -Apply $Apply
            if ($Apply -eq $True) {

                $Strt = Get-Date
                #Apply the MSI remaining in the $Global:DrvInfo
                Install-MSI
                $End = Get-Date
                $Span = New-TimeSpan -Start $Strt -End $End
                $Min = $Span.Minutes
                $Sec = $Span.Seconds
                
                Write-Host "Installed in $Min Min and $Sec Seconds"
            } else {
                if ($Expand -eq $True) {

                    $Strt = Get-Date
                    #Expand the MSI remaining in the $Global:DrvInfo
                    Expand-MSI
                    $End = Get-Date
                    $Span = New-TimeSpan -Start $Strt -End $End
                    $Min = $Span.Minutes
                    $Sec = $Span.Seconds
                    
                    Write-verbose "MSI Expended in $Min Min and $Sec Seconds"
                }
            }
    
        }

        return $True
    }

    end {
         Write-Verbose "End procesing Import-SurfaceDrivers"  
    }
}

$ModuleName = "Import-SurfaceDriver"
Write-Verbose "Loading $ModuleName Module"

Export-ModuleMember -Function Import-SurfaceDrivers,Import-SurfaceDB,Import-Config

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Verbose "Module $ModuleName Unloaded" }




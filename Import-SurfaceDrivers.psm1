function Get-LocalDriversInfo {
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
        [Alias('RootRepo')]
        [string]$RPath = '.\Repo'
        ,
        [Alias('Model')]
        [string]$DrvModel
        ,
        [Alias('OSTarget')]
        [string]$OVers
    )
    
    begin {
         Write-Verbose "Begin procesing Get-LocalDriversInfo"  
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
   }
  
    process {

        try {

            foreach ($DrvHashT in $Global:DrvInfo) {
                $OSSub = $DrvHashT['OSVersion']
                $LocalPathDir = "$RPath\$DrvModel\$OSSub"
                write-verbose "Testing if $LocalPathDir exist"
                If(!(test-path $LocalPathDir))
                    {
                        write-verbose "Create $LocalPathDir directory"
                        New-Item -ItemType Directory -Force -Path $LocalPathDir | out-null
                    }
                
                $FileName = $DrvHashT['FileName']
                write-verbose "Check for $FileName"
                $LocalPathFile = "$LocalPathDir\$FileName"
                $DrvHashT.Add('LPath',$LocalPathDir)

                If(!(test-path $LocalPathFile))
                    {
                        $DrvHashT.Add('LatestPresent',"N")
                        write-host "$FileName is missing"
                    }
                else
                    {
                        $DrvHashT.Add('LatestPresent',"Y")
                        write-Verbose "$FileName is already downloaded"                        
                        write-host "Driver for $OSSub found localy"                        
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
    <#
    .SYNOPSIS
    Check Online what are the available Driver Packages for a given Model
    and optionally ask for a OS Target version Package
    .DESCRIPTION
    Retun an Array containing hashtable containing OS Target and Download URL
    No MSI are downloaded from the function
    .EXAMPLE
    Get-RemoteDriversInfo "Surface Book"
    .EXAMPLE
    Get-RemoteDriversInfo -Model "Surface Pro4"
    .EXAMPLE
    Get-RemoteDriversInfo -Model "Surface Pro LTE" -OSTarget 1703
    .PARAMETER DrvModel
    Surface Model Name
    .PARAMETER OSTarget
    Optional - OS Version targeted by the driver package
    #>
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
        Write-Verbose "Begin procesing Get-RemoteDriversInfo($DrvModel,$Overs)"  
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
    }
  
    process {

        try {

            [System.Collections.ArrayList]$CurLst = @()
            [System.Collections.ArrayList]$FoundDrvLst = @()
            
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

            $DrvPage = Invoke-WebRequest -Uri $urldrv -UseBasicParsing

            foreach ($link in $DrvPage.Links) {

                $href = $link.href
                if ($href -ne $null) {
                    if ($href -like "*win10*.msi" ) {
                        if ($href.tolower() -notin $CurLst) {

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
                                        $ret = $FoundDrvLst.Add($FoundDrvHT)
                                        write-debug "Remote info for $VFound Added"
                                    }
                                }
                                else {
                                    $FoundDrvHT = @{} 
                                    $ret = $FoundDrvHT.Add("OSVersion",$VFound)
                                    $ret = $FoundDrvHT.Add("Link",$href.tolower())
                                    $ret = $FoundDrvHT.Add("FileName",$DrvPrsLst[$DrvPrsLst.count-1])
                                    $ret = $FoundDrvLst.Add($FoundDrvHT)
                                    write-debug "Remote info for $VFound Added"
                            }
                            $ret = $CurLst.Add($VFound)
                        }
                    }    
                }
            }
        }
    }
    catch [System.Exception] {
        Write-Host -ForegroundColor Red $_;
        return $false
    }

    return $FoundDrvLst
    }

    end {
         Write-Verbose "End procesing Get-RemoteDriversInfo"  
    }
}
function Import-SurfaceDB
    {

        $DBFileName = ".\ModelsDB.xml"
        If(test-path $DBFileName) {
            [XML]$ModelDBFile = Get-Content $DBFileName
        }
        else {
             Write-Verbose "Warning Model DB File not found"  

        }
    
        $ModelsHT = @{}                                 # empty models hashtable   

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

    $ConfFileName = ".\Config.xml"
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
    <#
    .SYNOPSIS
    Check and Set the sub folders directories that will hold the Surface Drivers
    .DESCRIPTION
    The function will check if the directories exists and if not it will create them
    .EXAMPLE
    Set-DriverRepo -Root '.\MyRoot' -SubFolders ('Surface Book','Surface Book 2')
    To set a Sub Folders structure with a root MyRoot and holding 2 Subfolders for 2 models of Surface
    .EXAMPLE
    Set-DriverRepo
    To user the function with the defaults
    .PARAMETER Root
    The Root of the Drivers Repo
    .PARAMETER SubFolders
    The list of Subfolders under the root. Each one will match a Surface Model
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0,mandatory=$true)]
        [Alias('Root')]
        [string]$RootRepo
        ,
        [Alias('SubFolders')]
        [string[]]$SubFolder = ('Surface Book','Surface Book 2','Surface Pro','Surface Laptop','Surface Studio','Surface Pro4','Surface Pro3')
    )
  
    begin {
        Write-Verbose "Begin procesing Drivers Repo"  
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
        [string]$targetFile
    )
    
    begin {
         Write-Verbose "Begin procesing Get-MSIFile"  
    }
  
    process {

        try {

            $uri = New-Object "System.Uri" "$url" 
            $request = [System.Net.HttpWebRequest]::Create($uri) 
            $request.set_Timeout(15000) #15 second timeout 
            $response = $request.GetResponse() 
            $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024) 
            $responseStream = $response.GetResponseStream() 
            $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create 
            $buffer = new-object byte[] 10KB 
            $count = $responseStream.Read($buffer,0,$buffer.length) 
            $downloadedBytes = $count 

            while ($count -gt 0) 
            { 
                $Downloaded = [System.Math]::Floor($downloadedBytes/1024)
                $PercentD = [math]::Round((100/$totalLength)*$Downloaded)
                if ($PercentD -gt 100) {$PercentD = 100}
                Write-Progress -Activity "Download MSI" -Status "$Downloaded K Downloaded:" -PercentComplete $PercentD
                $targetStream.Write($buffer, 0, $count) 
                $count = $responseStream.Read($buffer,0,$buffer.length) 
                $downloadedBytes = $downloadedBytes + $count 
            } 
            $targetStream.Flush()
            $targetStream.Close() 
            $targetStream.Dispose() 
            $responseStream.Dispose() 

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
        [Alias('Apply')]
        [Boolean]$ApplyDRv = $false
    )
    
    begin {
         Write-Verbose "Begin procesing Get-SurfaceDriver"  
    }
  
    process {

        try {

            foreach ($DrvHashT in $Global:DrvInfo) {
                if ($DrvHashT["LatestPresent"] -eq "N") {
                    $Link = $DrvHashT['Link']
                    $FileName = $DrvHashT['FileName']
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
                    $TagetLPath = "$LPath\$FileName"
                    Write-Host "Start Downloading .................... $FileName"
                    Get-MSIFile -Link $Link -LPath $TagetLPath
                    $End = Get-Date
                    $Span = New-TimeSpan -Start $Strt -End $End
                    $Min = $Span.Minutes
                    $Sec = $Span.Seconds

                    Write-Host "Downloaded in $Min Min and $Sec Seconds"
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
        [Alias('Model')]
        [string]$SurfaceModel
        ,
        [Alias('OSTarget')]
        [string]$OVers
        ,
        [Alias('Root')]
        [string]$RootRepo = '.\Repo'   
        ,
        [Alias('Apply')]
        [Boolean]$ApplyDRv = $false
        ,
        [Alias('Force')]
        [Boolean]$ForceApplyDRv = $false
        ,
        [Alias('CheckOnly')]
        [Boolean]$CheckOnlyDrv = $false  
    )

    begin {
        Write-Verbose "Begin procesing Import-SurfaceDrivers"
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
        $DefaultFromConfigFile = Import-Config
        $status = $false
    }
  
    process {

        if ($SurfaceModel -eq "") {
            if ('SurfaceModel' -in $DefaultFromConfigFile.Keys) {
                Write-Verbose "Getting SurfaceModel from Defaults in Config file"  
                $SurfaceModel = $DefaultFromConfigFile['SurfaceModel']
            }
            else {
                Write-Host -ForegroundColor Red "Surface Model need to be specified in Input or in the Config file"
                return $false
            }
        }

        if ($SurfaceModel.ToLower() -NotIn $SurfModelHT.keys) {
            Write-Host -ForegroundColor Red "Surface Model $SurfaceModel not supported by this tool"
            return $false
        }

        Write-Host "Check Drivers Repo for $SurfaceModel"

        If (Set-DriverRepo -RootRepo $RootRepo -SubFolders $SurfaceModel) {
             Write-Verbose "Drivers Repo Checked and Set"  
        }
        else {
             Write-Verbose "Error while checking Drivers Repo"
             return $false
        }

        if ($OVers -ne "") {
            $Global:DrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel -OSTarget $OVers
            if ($Global:DrvInfo -eq $null) {
                Write-Host -ForegroundColor Red "   Drivers not found for $OVers"
                return $false
            }
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel -OSTarget $OVers
        }
        else {
            $Global:DrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel
        }

        if ($CheckOnlyDrv -ne $True) {

            Get-SurfaceDriver -Apply $ApplyDRv

        }
        return $True
    }

    end {
         Write-Verbose "End procesing Import-SurfaceDrivers"  
    }
}
function Squel {
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
        [Alias('param1')]
        [string]$Param1 = 'initValue'
        ,
        [Alias('param2')]
        [int]$param2 = 1
    )
    
    begin {
         Write-Verbose "Begin procesing FunctionName"  
    }
  
    process {

        try {

            #Something

        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
         Write-Verbose "End procesing FunctionName"  
    }
}

$ModuleName = "Import-SurfaceDriver"
Write-Verbose "Loading $ModuleName Module"

Export-ModuleMember -Function Import-SurfaceDrivers,Import-SurfaceDB,Import-Config

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Verbose "Module $ModuleName Unloaded" }




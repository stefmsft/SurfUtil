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
         Write-Verbose "Begin procesing Get-LocalDriversInfo(Repo=$RPath,Model=$DrvModel,OSVersion=$OVers)"  
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
        Write-Verbose "Begin procesing Get-RemoteDriversInfo(Model=$DrvModel,OSVersion=$Overs)"  
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
    <#
    .SYNOPSIS
    Check and Set the sub folders directories that will hold the Surface Drivers
    .DESCRIPTION
    The function will check if the directories exists and if not it will create them
    .EXAMPLE
    Set-DriverRepo -Root '.\MyRoot' -Model ('Surface Book','Surface Book 2')
    To set a Sub Folders structure with a root MyRoot and holding 2 Subfolders for 2 models of Surface
    .EXAMPLE
    Set-DriverRepo
    To user the function with the defaults
    .PARAMETER Root
    The Root of the Drivers Repo
    .PARAMETER Model
    The list of Subfolders under the root. Each one will match a Surface Model
    #>
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
         Write-Verbose "Begin procesing Get-MSIFile(Link=$url,LPath=$TargetPath,File=$targetFile)"  
    }
  
    process {

        try {

            $FullTP = resolve-path $TargetPath
            $FulltargetFile = "$FullTP\$TargetFile"
            write-host $FulltargetFile

            $uri = New-Object "System.Uri" "$url" 
            $request = [System.Net.HttpWebRequest]::Create($uri) 
            $request.set_Timeout(15000) #15 second timeout 
            $response = $request.GetResponse()
            $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024) 
            $responseStream = $response.GetResponseStream() 
            $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $FulltargetFile, Create 
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
         Write-Verbose "Begin procesing Get-SurfaceDriver(Apply=$ApplyDrv)"  
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
                    #$TagetLPath = "$LPath\$FileName"
                    Write-Host "Start Downloading .................... $FileName"
                    Get-MSIFile -Link $Link -LPath $LPath -File $FileName
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
    Function to import Surface Drivers
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
        [Boolean]$ApplyDRv = $False
        ,
        [Alias('Force')]
        [Boolean]$ForceApplyDRv = $False
        ,
        [Alias('CheckOnly')]
        [Boolean]$CheckOnlyDrv = $False  
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

        if ($OVers -ne "") {
            Write-Host "Check Drivers Repo for $SurfaceModel for Windows $OVers"
        }
        else {
            Write-Host "Check Drivers Repo for $SurfaceModel"            
        }

        If (Set-DriverRepo -RootRepo $RootRepo -Model $SurfaceModel) {
             Write-Verbose "Drivers Repo Checked and Set"  
        }
        else {
             Write-Verbose "Error while checking Drivers Repo"
             return $false
        }

        if ($OVers -ne "") {
            $Global:DrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel -OSTarget $OVers
            if ($Global:DrvInfo -eq $null) {
                Write-Host ">>>   Drivers not found for $OVers"
                return $false
            }
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel -OSTarget $OVers
        }
        else {
            $Global:DrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel
            if ($Global:DrvInfo -eq $null) {
                Write-Host -ForegroundColor Red "   No Drivers found "
                return $false
            }
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel
        }

        if ($ApplyDRv -eq $True) {
            #Trim $Global:DrvInfo to the latest Driver only
            $intver=0
            $i=0
            #Search the highest Version number
            foreach ($DrvHashT in $Global:DrvInfo) {
                $curintver = [convert]::ToInt32($DrvHashT['OSVersion'], 10)
                if($curintver -gt $intver) {
                    $HT = $DrvHashT
                    $intver = $curintver
                }
                $i = $i + 1
            }
            if ($intver -ne 0 ) {
                write-verbose "Found $intver as the latest version available"
                $Global:DrvInfo = $HT
            }
        }

        if ($CheckOnlyDrv -ne $True) {

            Get-SurfaceDriver -Apply $ApplyDRv
            if ($ApplyDRv -eq $True) {

                $Strt = Get-Date
                #Apply the MSI remaining in the $Global:DrvInfo
                Install-MSI
                $End = Get-Date
                $Span = New-TimeSpan -Start $Strt -End $End
                $Min = $Span.Minutes
                $Sec = $Span.Seconds

                Write-Host "Installed in $Min Min and $Sec Seconds"
            }
    
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




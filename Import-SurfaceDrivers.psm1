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

            foreach ($DrvHashT in $Global:RemDrvInfo) {
                $OSSub = $DrvHashT['OSVersion']
                $LocalPathDir = "$RPath\$DrvModel\$OSSub"
                write-verbose "Testing if $LocalPathDir exist"
                If(!(test-path $LocalPathDir))
                    {
                        write-verbose "Create $LocalPathDir directory"
                        New-Item -ItemType Directory -Force -Path $LocalPathDir | out-null
                    }
                
                $Link = $DrvHashT['Link']
                $SplitedLink = $Link.tolower() -split '/'
                $FileName = $SplitedLink[$SplitedLink.count-1]
                write-verbose "Check for $FileName"
                $LocalPathFile = "$LocalPathDir\$FileName"

                If(!(test-path $LocalPathFile))
                    {
                        write-host "$FileName is missing"
                    }
                }
        return $true
        }
        catch [System.Exception] {
            Write-Error $_.Exception.Message;
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
        Write-Verbose "Begin procesing Get-RemoteDriversInfo"  
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
    }
  
    process {

        try {

            [System.Collections.ArrayList]$CurLst = @()
            [System.Collections.ArrayList]$FoundDrvLst = @()
            
            if ($DrvModel -ne "") {
                $urldrv = $SurfModelHT[$DrvModel.tolower()]
                if ($urldrv -eq $null) {
                    Write-Error "Unknown Surface Model for the script : [$DrvModel]"
                    return $false
                }
            }

            if (($OVers -ne $null) -and ($OVers -ne "")) {
                $InternalR = $OSReleaseHT[$OVers.tolower()]
                if ($InternalR -eq $null) {
                    Write-Error "Unknown OS Release for the script : [$OVers]"
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

                            if ($OSReleaseHT.containsValue($InternalVFound)) {
                                $VFound = $OSReleaseHT.Keys | ForEach-Object { if ($OSReleaseHT.Item($_) -eq $InternalVFound ) {$_} }
                            }
                            else {
                                $VFound = "1507"
                            }

                            $ret = $CurLst.Add($href.tolower())

                             Write-Verbose "[$ret]:$href"   
                             Write-Verbose "Found OS Version is $VFound"  
                            if ($OVers -ne "") {
                                 Write-Verbose "Asked OS Version is $OVers"  
                                if ($Overs -eq $VFound) {
                                    $FoundDrvHT = @{} 
                                    $ret = $FoundDrvHT.Add("OSVersion",$VFound)
                                    $ret = $FoundDrvHT.Add("Link",$href.tolower())
                                    $ret = $FoundDrvLst.Add($FoundDrvHT)
                                }
                            }
                            else {
                                $FoundDrvHT = @{} 
                                $ret = $FoundDrvHT.Add("OSVersion",$VFound)
                                $ret = $FoundDrvHT.Add("Link",$href.tolower())
                                $ret = $FoundDrvLst.Add($FoundDrvHT)
                            }
                        }
                    }    
                }
            }
        }
        catch [System.Exception] {
            Write-Error $_;
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

#    Write-Debug $ModelsHT.Keys
#    Write-Debug $ModelsHT.Values
#    Write-Debug $OSHT.Keys
#    Write-Debug $OSHT.Values

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

#    Write-Debug $Defaults.Keys
#    Write-Debug $Defaults.Values

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
            Write-Error $_.Exception.Message;
            return $False
        }

        return $True

      }

      end {
         Write-Verbose "End procesing Driver Repo"  
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
                Write-Error "Surface Model need to be specified in Input or in the Config file"
                return $false
            }
        }

        if ($SurfaceModel.ToLower() -NotIn $SurfModelHT.keys) {
            Write-Error "Surface Model $SurfaceModel not supported by this tool"
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
            $Global:RemDrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel -OSTarget $OVers
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel -OSTarget $OVers
        }
        else {
            $Global:RemDrvInfo = Get-RemoteDriversInfo -DrvModel $SurfaceModel
            $status = Get-LocalDriversInfo -RootRepo $RootRepo -DrvModel $SurfaceModel
        }

#        write-debug $Global:RemDrvInfo.keys 
#        write-debug $Global:RemDrvInfo.Values 
        return $status
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
            Write-Error $_.Exception.Message;
            return $False
        }

        return $True
    }

    end {
         Write-Verbose "End procesing FunctionName"  
    }
}



$ModuleName = "Import-SurfaceDriver"
Write-Host "Loading $ModuleName Module"

Export-ModuleMember -Function Import-SurfaceDrivers

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Host "Module $ModuleName Unloaded" }




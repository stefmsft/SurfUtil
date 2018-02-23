function Import-SurfaceDB
    {

        $DBFileName = ".\ModelsDB.xml"
        If(test-path $DBFileName) {
            [XML]$ModelDBFile = Get-Content $DBFileName
        }
        else {
            Write-Verbose "Warning Model DB File not found" -Verbose

        }
    
        $ModelsHT = @{}   # empty models hashtable

        foreach ($Child in $ModelDBFile.ModelsDB.SurfacesModels.ChildNodes ) {

                $ModelsHT.Add($Child.ID,$Child.Drivers.url)
        }

        $OSHT = @{}   # empty models hashtable

        foreach ($Child in $ModelDBFile.ModelsDB.OSRelease.ChildNodes ) {

                $OSHT.Add($Child.Name,$Child.ReleaseCode)
        }

    Write-Debug $ModelsHT.Keys
    Write-Debug $ModelsHT.Values
    Write-Debug $OSHT.Keys
    Write-Debug $OSHT.Values

    $CtxData = ($ModelsHT,$OSHT)

    return $CtxData
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
        [string]$SurfaceModel = 'Surface Pro'
    )

    begin {
        Write-Verbose "Begin procesing Import-SurfaceDrivers" -Verbose
        ($SurfModelHT,$OSReleaseHT) = Import-SurfaceDB
    }
  
    process {

        Write-Host "Check Drivers Repo for $SurfaceModel"
        
#        $urldrv = $SurfModelHT[$SurfaceModel]
#        Write-Debug "$SurfaceModel, Driver url : $urldrv"

        If (Set-DriverRepo -SubFolders $SurfaceModel) {
            Write-Verbose "Drivers Repo Checked and Set" -Verbose
        }
        else {
            Write-Verbose "Error while checking Drivers Repo" -Verbose
        }

        return $True
    }

    end {
        Write-Verbose "End procesing Import-SurfaceDrivers" -Verbose
    }
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
        [Alias('Root')]
        [string]$RootRepo = '.\Repo'
        ,
        [Alias('SubFolders')]
        [string[]]$SubFolder = ('Surface Book','Surface Book 2','Surface Pro','Surface Laptop','Surface Studio','Surface Pro4','Surface Pro3')
    )
  
    begin {
        Write-Verbose "Begin procesing Drivers Repo" -Verbose
    }
  
    process {
  
        Write-Debug "Verify $RootRepo"
        try {

            If(!(test-path $RootRepo))
            {
                Write-Debug "Create $RootRepo"
                New-Item -ItemType Directory -Force -Path $RootRepo
            }

            foreach ($s in $SubFolder) {
                Write-Debug "Processing $s"
                $TstSub = "$RootRepo\$s"
                Write-Debug "Verify Directory $TstSub"
                If(!(test-path $TstSub))
                {
                    New-Item -ItemType Directory -Force -Path $TstSub
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
        Write-Verbose "End procesing Driver Repo" -Verbose
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
        Write-Verbose "Begin procesing FunctionName" -Verbose
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
        Write-Verbose "End procesing FunctionName" -Verbose
    }
}



$ModuleName = "Import-SurfaceDriver"
Write-Host "Loading $ModuleName Module"

Export-ModuleMember -Function Import-SurfaceDrivers

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Host "Module $ModuleName Unloaded" }
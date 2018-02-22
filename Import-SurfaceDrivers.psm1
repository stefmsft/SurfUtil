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
    begin {
        Write-Verbose "Begin procesing Import-SurfaceDrivers" -Verbose
    }
  
    process {

        If (Set-DriverRepo) {
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
        Write-Verbose "Begin procesing Driver Repo" -Verbose
    }
  
    process {
  
        Write-Debug "Processing $RootRepo"
        try {

            If(!(test-path $RootRepo))
            {
                New-Item -ItemType Directory -Force -Path $RootRepo
            }
      
            foreach ($s in $SubFolder) {
                Write-Debug "Processing $s"
                $TstSub = "$RootRepo\$s"
                Write-Debug "Create Directory $TstSub"
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

$ModuleName = "Import-SurfaceDriver"
Write-Host "Loading $ModuleName Module"

Export-ModuleMember -Function Import-SurfaceDrivers

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Host "Module $ModuleName Unloaded" }
function New-IsoFile  {  
  <#  
   .Synopsis  
    Creates a new .iso file  
   .Description  
    The New-IsoFile cmdlet creates a new .iso file containing content from chosen folders  
   .Example  
    New-IsoFile "c:\tools","c:Downloads\utils"  
    This command creates a .iso file in $env:temp folder (default location) that contains c:\tools and c:\downloads\utils folders. The folders themselves are included at the root of the .iso image.  
   .Example 
    New-IsoFile -FromClipboard -Verbose 
    Before running this command, select and copy (Ctrl-C) files/folders in Explorer first.  
   .Example  
    dir c:\WinPE | New-IsoFile -Path c:\temp\WinPE.iso -BootFile "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\efisys.bin" -Media DVDPLUSR -Title "WinPE" 
    This command creates a bootable .iso file containing the content from c:\WinPE folder, but the folder itself isn't included. Boot file etfsboot.com can be found in Windows ADK. Refer to IMAPI_MEDIA_PHYSICAL_TYPE enumeration for possible media types: http://msdn.microsoft.com/en-us/library/windows/desktop/aa366217(v=vs.85).aspx  
   .Notes 
    NAME:  New-IsoFile  
    AUTHOR: Chris Wu 
    LASTEDIT: 03/23/2016 14:46:50  
 #>  
  
  [CmdletBinding(DefaultParameterSetName='Source')]Param( 
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]$Source,  
    [parameter(Position=2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch]$Force, 
    [parameter(ParameterSetName='Clipboard')][switch]$FromClipboard 
  ) 
 
  Begin {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @' 
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
  
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
  
    if ($BootFile) { 
        if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
        ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
        $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
        ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
      } 
   
      $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
   
      Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))" 
      ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
    
      if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
    }  
   
    Process { 
      if($FromClipboard) { 
        if($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
        $Source = Get-Clipboard -Format FileDropList 
      } 
   
      foreach($item in $Source) { 
        if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
          $item = Get-Item -LiteralPath $item 
        } 
   
        if($item) { 
          Write-Verbose -Message "Adding item to the target image: $($item.FullName)" 
          try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
        } 
      } 
    } 
   
    End {  
      if ($Boot) { $Image.BootImageOptions=$Boot }  
      $Result = $Image.CreateResultImage()  
      [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
      Write-Verbose -Message "Target image ($($Target.FullName)) has been created" 
      $Target 
    } 
  } 

  function Set-USBKey {
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
        [Alias('Drive')]
        [string]$Drv
        ,
        [Alias('SrcISO')]
        [string]$Src
        ,
        [Alias('Model')]
        [string]$SurfaceModel
        ,
        [Alias('TargOS')]
        [string]$TargetedOS
        ,
        [Alias('DrvRepo')]
        [string]$DriverRepo
        ,
        [Alias('Sku')]
        [string]$TSku="Windows 10 Pro"
    )
    
    begin {
         Write-Verbose "Begin procesing Set-USBKey($Drv,$Src,$SurfaceModel,$TargetedOS,$DriverRepo,$TSku)"
    }
  
    process {

        try {

            #Check on the Destination
            $TargetDrv = Get-WmiObject win32_volume|where-object {$_.driveletter -match "$Drv"}
            $TargetSize = [int32]($TargetDrv.Capacity / 1GB)
            $TargetLabel = $TargetDrv.label
            write-verbose "Target Drive size is $TargetSize GB"
            write-verbose "Target Drive label is $TargetLabel"

            if (($TargetSize -lt 4) -Or ($TargetSize -gt 50)) {
                Write-Host -ForegroundColor Red "Use a USB Drive with capacity between 4 and 50 Go";
                return $False                
            }

            $message  = "If you agree, the external drive labeled [$TargetLabel] will be formated"
            $question = 'Are you sure you want to proceed?'

            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

            $decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
            if ($decision -eq 1) {

                Write-Host 'Action canceled ...'
                return $False
                
            }

            #Check on the Source
            If(!(test-path $SrcISO)) {
                write-verbose "$SrcISO directory doesn't exist"
                return $False
            }

            $ISOFilter = "*windows_10*_$TargetedOS*"
            Write-Verbose "Filtering ISO directory on $ISOFilter"
            $ISODir = get-childitem $SrcISO
            $ListMatchingISO = $ISODir | where-object {$_.Name -like $ISOFilter}
            if ($ListMatchingISO.count -ne 0) {
                foreach ($ISOfile in $ListMatchingISO) {
                    write-verbose "Found $ISOFile"
                    $SrcISO = "$SrcISO\$ISOfile"
                }
            }

            $MountedLetter = ""
            $MountedISO = Mount-DiskImage -ImagePath $SrcISO -PassThru
            $MountedLetter = ($MountedISO | Get-Volume).DriveLetter
            $IsoMounted = $true
            Write-Verbose "ISO mounted on $MountedLetter"

            $TargetInstalWim = $MountedLetter+":\Sources\install.wim"
            Write-Verbose "Checking $TargetInstalWim"
            if (!(Test-Path $TargetInstalWim)) {
                Write-Host -ForegroundColor Red "The mounted ISO doesn't hold a valid Wim to use";
                Write-Verbose "Dismounting $SrcISO"
                DisMount-DiskImage -ImagePath $SrcISO
                $IsoMounted = $false
                return $False
            }

            # Verify Wim Tmp held directory
            $TmpTDir = "$env:TMP\Wimdir\"
            write-verbose "Test if $TmpTDir exist"
            If(!(test-path $TmpTDir)) {
                write-verbose "Create $TmpTDir directory"
                New-Item -ItemType Directory -Force -Path $TmpTDir | out-null
            } else {
                $Dir = get-childitem $TmpTDir
                $ListExistingWim = $Dir | where-object {$_.extension -eq ".wim"}
                if ($ListExistingWim.count -ne 0) {
                    foreach ($wimfile in $ListExistingWim) {
                        $WimToRemove = "$TmpTDir$wimfile"
                        write-verbose "Removing $wimfile in $TmpTDir directory"
                        Remove-Item $WimToRemove -Force
                    }
                }
            }

            # Verify Mounting directory
            $MntDir = "$env:TMP\Wimdir\Mounted"
            write-verbose "Test if $MntDir exist"
            If(!(test-path $MntDir)) {
                write-verbose "Create $MntDir directory"
                New-Item -ItemType Directory -Force -Path $MntDir | out-null
            } else {
                write-verbose "Removing all file in $MntDir directory"
                $ListExistingFile = Get-ChildItem $MntDir
                if ($ListExistingFile.count -ne 0) {
                    $Ret = Dismount-WindowsImage -Path $MntDir -Discard
                    # Remove-Item $MntDir -Force -recurse | Out-Null
                }
            }

            
            # Verify Scratch directory
            $ScrtchDir = "$env:TMP\Wimdir\Scratch"
            write-verbose "Test if $ScrtchDir exist"
            If(!(test-path $ScrtchDir)) {
                write-verbose "Create $ScrtchDir directory"
                New-Item -ItemType Directory -Force -Path $ScrtchDir | out-null
            } else {
                write-verbose "Removing all file in $ScrtchDir directory"
                Remove-Item $ScrtchDir -Force -recurse | Out-Null
            }

            # proceed to format and copy all file except WIM
            write-host "The next steps are :"
            write-host "    1 - Format the target and copy windows files"
            write-host "    2 - Prepare the Wim File"
            write-host "    3 - Inject Surface Drivers"
            write-host "    4 - Apply latest Windows Update"
            write-host "    5 - Copy the wim to the Key"
            write-host ""
            write-host "Please, don't interrupt the script ...."

            [System.Collections.ArrayList]$JbLst = @()

######
# Job 1 Format target and copy windows file
######

            if ($Global:SkipJ1 -ne $true) {

                Write-Verbose "Firing Up Job 1 ..."
                $JobFormatAndCopy = Start-Job -Name "Job1" -ScriptBlock {

                    param ([String] $TDrv, [String] $MISO, [String] $SurfaceModel, [String] $TOS, [String] $v)

                        $VerbosePreference=$v 
                        write-verbose "Call Job 1 ($TDrv,$MISO,$SurfaceModel,$TOS)"
                
                        $Strt = Get-Date
                        write-host ":: Step 1"
                        write-host "========="
                        write-host "Formatting drive $TDrv ....."
    
                        Format-Volume -DriveLetter $TDrv -FileSystem FAT32 -NewFileSystemLabel "BMR $TOS"
    
                        write-host "Copying files....."
                        Copy-Item -Path ($MISO+":\*") -Destination ($TDrv+":\") -Exclude oinstall.wim, install.wim, *.swm -Force -Recurse
    
                        #Write tag file on completion
                        $ModelTag = $SurfaceModel.Replace(" ","")
                        $TagFileName = $TDrv+":\$ModelTag-$TOS.tag"
                        Write-Host "Tagging with file $TagFileName"
    
                        New-Item $TagFileName -type file
    
                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "Format/Copy in $Min Min and $Sec Seconds"
    
                } -ArgumentList $Drv, $MountedLetter, $SurfaceModel, $TargetedOS, $VerbosePreference
                $ret = $JbLst.Add($JobFormatAndCopy) 

            } else {

            Write-verbose "Job 1 Skipped"

        }

######
# Job 2 Prepare the Wim File
######
            
            if ($Global:SkipJ2 -ne $true) {

                Write-Verbose "Firing Up Job 2 ..."
                $JobPrepareWIM = Start-Job -Name "Job2" -ScriptBlock {

                        param ([String] $SrcWimLetter, [String] $TDir, [String] $TargetSKU, [String] $v)
    
                        $VerbosePreference=$v
                        write-verbose "Call Job 2 ($SrcWimLetter,$TDir,$TargetSKU)"

                        $Strt = Get-Date
                        write-host ":: Step 2"
                        write-host "========="
        
                        [String]$SrcWimFile = (Join-Path -Path "$SrcWimLetter" -ChildPath '\sources\install.wim')
                        [String]$DstWimFile = (Join-Path -Path "$TDir" -ChildPath '\install.wim')
                        write-host "Copying $WimToRemove install.wim"
                        copy-item $SrcWimFile -Destination $DstWimFile
                        Set-ItemProperty $DstWimFile -name IsReadOnly -value $false
                        $skul = Get-WindowsImage -ImagePath $DstWimFile
                        $sku = $skul | where-object {$_.imagename.tolower() -eq $TargetSKU.tolower()}
                        $indx = $sku.imageindex
                        Write-verbose "Found [$TargetSKU] at index $indx"

                        if ($sku -eq $null) {

                            $indx = -1
                            write-host "Job 2 : Operation Failed"
                            write-host "         The target SKU is not present in the wim file"

                        }

        
                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "Wim Prepared in $Min Min and $Sec Seconds"
        
                        return $indx
                        
                } -ArgumentList ($MountedLetter+":"), $TmpTDir, $TSku, $VerbosePreference
                $ret = $JbLst.Add($JobPrepareWIM) 
    
            } else {

            Write-verbose "Job 2 Skipped"

        }

######
# Job 3 Expand the Surface drivers for the targeted BMR
######

            if ($Global:SkipJ3 -ne $true) {

                $localp = (Get-Item -Path ".\" -Verbose).FullName
                $mp = "$localp\Import-SurfaceDrivers.psm1"

                Write-Verbose "Firing Up Job 3 ..."
                $JobGetExpandDrv = Start-Job -Name "Job3" -ScriptBlock {

                    param ([String] $Model, [String] $os, [String] $LocalRepoPathDir, [String] $IDMPath, [String] $v)

                        $VerbosePreference=$v
                        write-verbose "Call Job 3 ($Model,$os,$LocalRepoPathDir,$IDMPath)"
                        $Strt = Get-Date
                        write-host ":: Step 3"
                        write-host "========="

                        $ret = import-module $IDMPath
        
                        $ret = Import-SurfaceDrivers -Model $Model -OSTarget $os -Root $LocalRepoPathDir -Expand $True
        
                        if ($ret -eq $false) {
                            write-host "Step 3 : Operation Failed"
                            write-host "         Drivers expand from msi failed"
                        }
            
                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "Drivers gathered in $Min Min and $Sec Seconds"    
        
                } -ArgumentList $SurfaceModel, $TargetedOS, $DriverRepo, $mp, $VerbosePreference

                $ret = $JbLst.Add($JobGetExpandDrv) 
                
            } else {

                Write-verbose "Job 3 Skipped"

            }

            Wait-Job $JbLst | Out-Null

            if ($VerbosePreference -eq "Continue") {

                foreach ($j in $JbLst){
                    Receive-Job $jget-it    
                }
            }

            $WimIndx = -1

            foreach ($j in $JbLst){
                # Gather Wim Index if job 2 was fired
                if ($j.Name -eq "Job2") {
                    $WimIndx = [convert]::ToInt32($j.ChildJobs[0].Output[0],10)
                    write-verbose "Job 2 returned $WimIndx"
                }
            }

            if ($WimIndx -ge 0) {

                ## Mount Wim
                [String]$MntWimFile = (Join-Path -Path $TmpTDir -ChildPath 'install.wim')
                [String]$LogPath = (Join-Path -Path $TmpTDir -ChildPath 'DISM.log')
                Write-Host "Mounting Image..."
                $Ret = Mount-WindowsImage -ImagePath $MntWimFile -Index $WimIndx -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath
                $WimMounted = $true

                ## Get Wim infos
                $ret = Get-WindowsImage -Mounted -ScratchDirectory $ScrtchDir -LogPath $LogPath | Out-String
                Write-verbose "Mounted Image Info:`n $ret"

                $Discard = $False

                ## Inject Drivers
                [String]$DrvExpandRoot = (Join-Path -Path $env:TMP -ChildPath '\ExpandedMSIDir\SurfacePlatformInstaller')
                If(!(test-path $DrvExpandRoot)) {
                    write-verbose "Drivers not present ... Job 3 might have failed"
                    $Discard = $true
                } else {
                    Add-WindowsDriver -Path $MntDir -Driver $DrvExpandRoot -Recurse -LogPath $LogPath | Out-Null
                    Write-Host "Drivers injected in the new Wim"
                }

                ## Unmount and save servicing changes to the image
                if ($Discard -eq $true) {
                    Write-Host "Discard Changes and Dismounting Image..."
                    $Ret = Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Discard
                } else {
                    Write-Host "Committing Changes and Dismounting Image..."
                    $Ret = Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Save
                }
                $WimMounted = $false

                #Export the image to a more compact one
                #TBD generate a nice name for the wim
                [String]$ExptWimFile = (Join-Path -Path $TmpTDir -ChildPath 'install_serviced.wim')
                $Ret = Export-WindowsImage -SourceImagePath $MntWimFile -SourceIndex $WimIndx -DestinationImagePath $ExptWimFile -DestinationName "BMR" -ScratchDirectory $ScrtchDir -LogPath $LogPath

                [String]$FinalSwmFile = (Join-Path -Path ($MountedLetter + ":") -ChildPath '\sources\install.swm')
                $ret = Split-WindowsImage -FileSize 3500 -ImagePath $ExptWimFile -SplitImagePath $FinalSwmFile -CheckIntegrity -ScratchDirectory $ScrtchDir

            }

            Write-Verbose "Dismounting $SrcISO"
            DisMount-DiskImage -ImagePath $SrcISO
            $IsoMounted = $false

            # Generate ISO file
            #TBD Generate nice names 
            Get-ChildItem ($MountedLetter + ":\") | New-IsoFile -Media 'DVDPLUSR' -Title "BMR"

            return $true

        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }
        finally {
            # Check if the ISO need to be dismounted
            if ($IsoMounted -eq $true) {
                Write-Verbose "Dismounting ISO after Exception"
                DisMount-DiskImage -ImagePath $SrcISO
            }
            # Check if the Wim need to be dismounted
            if ($WimMounted -eq $true) {
                Write-Verbose "Dismounting Wim after Exception"
                Dismount-WindowsImage -Path $MntWimFile -ScratchDirectory $ScrtchDir -LogPath $LogPath -Discard
            }
            #Do some cleaning
            Get-Job | Remove-Job
            # Cleanup the misc Temp dir except we specify to keep them for debug purpose
            if ($Global:KeepExpandedDir -ne $true) {
                
                #TBD Remove Expanded dir content

            }

            if ($Global:KeepDriversFile -ne $true) {
                
                #TBD Remove SDrivers files

            }

            if ($Global:KeepWimDir -ne $true) {
                
                #TBD Remove WimDir

            }

        }
    }

    end {
         Write-Verbose "End procesing Set-USBKey"  
    }
}
function New-USBKey {
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
        [Alias('Drive')]
        [string]$DrvLetter
        ,
        [Alias('ISOPath')]
        [string]$SrcISO
        ,
        [Alias('DRVPath')]
        [string]$DrvRepoPath
        ,
        [Alias('Model')]
        [string]$SurfaceModel
        ,
        [Alias('OSV')]
        [string]$TargetedOS
    )
    
    begin {
        Write-Verbose "Begin procesing New-USBKey($DrvLetter,$SrcISO,$DrvRepoPath,$SurfaceModel,$TargetedOS)"
        $Global:SkipJ1 = $false
        $Global:SkipJ2 = $false
        $Global:SkipJ3 = $false
        $Global:KeepExpandedDir = $false
        $Global:KeepDriversFile = $false
        $Global:KeepWimDir = $false
    }
  
    process {

        try {

            #Display in verbose the Skip Jobs flags
            Write-Verbose "Skip Job 1 : $Global:SkipJ1"
            Write-Verbose "Skip Job 2 : $Global:SkipJ2"
            Write-Verbose "Skip Job 3 : $Global:SkipJ3"

            Write-Verbose "Verbosity : $VerbosePreference"

            Write-Verbose "Keep Expanded Dir : $Global:KeepExpandedDir"
            Write-Verbose "Keep SDrivers files : $Global:KeepDriversFile"
            Write-Verbose "Keep WimDir Dir : $Global:KeepWimDir"

            #Expand $SrcISO directory in full path string
            if ($SrcISO -ne "") {

                $SrcISO = (Get-Item -Path $SrcISO -Verbose).FullName
            
            }
            else {

                Write-Host -ForegroundColor Red "Source directory for ISOs is missing"
                return $False

            }

            #Expand $DrvRepoPath directory in full path string
            if ($DrvRepoPath -ne "") {

                $DrvRepoPath = (Get-Item -Path $DrvRepoPath -Verbose).FullName
            
            }
            else {

                Write-Host -ForegroundColor Red "Source directory for drivers repo is missing"
                return $False

            }
            
            Set-USBKey -Drive $DrvLetter -SrcISO $SrcISO -Model $SurfaceModel -TargOS $TargetedOS -DrvRepo $DrvRepoPath

        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
         Write-Verbose "End procesing New-USBKey"
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

$ModuleName = "Manage-BMR"
Write-Verbose "Loading $ModuleName Module"

Export-ModuleMember -Function New-USBKey,New-IsoFile 

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Verbose "Module $ModuleName Unloaded" }

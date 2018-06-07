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
        [Alias('MkIso')]
        [string]$MakeISO
        ,
        [Alias('Sku')]
        [string]$TSku
    )
    
    begin {
        Write-Verbose "Begin procesing Set-USBKey($Drv,$Src,$SurfaceModel,$TargetedOS,$DriverRepo,MakeISO=$MakeISO,$TSku)"
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
#            write-host "    4 - Apply latest Windows Update"
            write-host "    5 - Optimize and copy the wim to the Key"
            if ($MakeISO -eq $True) {write-host "    6 - Generate an ISO copy of your USB Key"}
            write-host ""
            write-host "Please, don't interrupt the script ...."

            $Strt = Get-Date
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
    
                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "Format/Copy in $Min Min and $Sec Seconds"
                        write-host "============="
                        write-host ":: End Step 1"
                        write-host "============="
    
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
                        write-host "============="
                        write-host ":: End Step 2"
                        write-host "============="
        
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
                $mp = "$localp\SurfUtil.psm1"

                Write-Verbose "Firing Up Job 3 ..."
                $JobGetExpandDrv = Start-Job -Name "Job3" -ScriptBlock {

                    param ([String] $Model, [String] $os, [String] $LocalRepoPathDir, [String] $IDMPath, [String] $v)

                        $VerbosePreference=$v
                        write-verbose "Call Job 3 ($Model,$os,$LocalRepoPathDir,$IDMPath)"
                        $Strt = Get-Date
                        write-host ":: Step 3"
                        write-host "========="

                        $ret = import-module $IDMPath
        
                        Write-Verbose "Calling : Import-SurfaceDrivers -Model $Model -WindowsVersion $os -RepoPath $LocalRepoPathDir -Expand $True"
                        $ret = Import-SurfaceDrivers -Model $Model -WindowsVersion $os -RepoPath $LocalRepoPathDir -Expand $True

                        #Drivers not found for the specific version on Windows. Let's try to find the latest Drivers available
                        if ($ret -eq $False) {

                            Write-Verbose "Calling : Import-SurfaceDrivers -Model $Model -RepoPath $LocalRepoPathDir -Expand $True"
                            $ret = Import-SurfaceDrivers -Model $Model -RepoPath $LocalRepoPathDir -Expand $True

                        }
        
                        if ($ret -eq $false) {
                            write-host "Step 3 : Operation Failed"
                            write-host "         Drivers expand from msi failed"
                        }
            
                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "Drivers gathered in $Min Min and $Sec Seconds"    
                        write-host "============="
                        write-host ":: End Step 3"
                        write-host "============="
        
                } -ArgumentList $SurfaceModel, $TargetedOS, $DriverRepo, $mp, $VerbosePreference

                $ret = $JbLst.Add($JobGetExpandDrv) 
                
            } else {

                Write-verbose "Job 3 Skipped"

            }

            Wait-Job $JbLst | Out-Null

            if ($VerbosePreference -eq "Continue") {

                foreach ($j in $JbLst){
                    Receive-Job $j    
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

                ## Preliminary test on the ExpandedMSIDir to avoid the useless mount of the wim after
                [String]$DrvExpandRoot = (Join-Path -Path $env:TMP -ChildPath '\ExpandedMSIDir\SurfacePlatformInstaller')

                If(test-path $DrvExpandRoot) {
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
                    Add-WindowsDriver -Path $MntDir -Driver $DrvExpandRoot -Recurse -LogPath $LogPath | Out-Null
                    Write-Host "Drivers injected in the new Wim"
                    write-verbose "Add the Driver Readme.txt to the USB Key"
                    $ReadMe = "$DrvExpandRoot\ReadMe.txt"
                    Copy-Item -Path $ReadMe -Destination ($Drv+":\ReadMe.txt")

                    ## Unmount and save servicing changes to the image
                    if ($Discard -eq $true) {
                        Write-Host "Discard Changes and Dismounting Image..."
                        $Ret = Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Discard
                    } else {
                        Write-Host "Committing Changes and Dismounting Image..."
                        $Ret = Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Save
                    }
                } else {

                    $Discard =$true

                }

                $WimMounted = $false

                if ($Discard -eq $true) {

                    $CompactName = "ERROR"
                    
                } else {

                    $CompactName = "BMR-" + $SurfaceModel.replace(" ","").toupper() + $TargetedOS
                    #Export the image to a more compact one
    
                    [String]$ExptWimFile = (Join-Path -Path $TmpTDir -ChildPath 'install_serviced.wim')
                    $Ret = Export-WindowsImage -SourceImagePath $MntWimFile -SourceIndex $WimIndx -DestinationImagePath $ExptWimFile -DestinationName $CompactName -ScratchDirectory $ScrtchDir -LogPath $LogPath
    
                    [String]$FinalSwmFile = (Join-Path -Path ($Drv + ":") -ChildPath '\sources\install.swm')
                    $ret = Split-WindowsImage -FileSize 3500 -ImagePath $ExptWimFile -SplitImagePath $FinalSwmFile -CheckIntegrity -ScratchDirectory $ScrtchDir
    
                }
            }

            Write-Verbose "Dismounting $SrcISO"
            DisMount-DiskImage -ImagePath $SrcISO
            $IsoMounted = $false

            # Generate ISO file
            #TBD Generate nice names 
            if (($MakeISO -eq $true) -and ($Discard -ne $true)) {
                
                $TCIsoName = ".\" + $CompactName + ".iso"
                Get-ChildItem ($Drv + ":\") | New-IsoFile -Media 'DVDPLUSR_DUALLAYER' -Title $CompactName -Path $TCIsoName -Force

            }

            #Write tag file on completion
            $TagFileName = $Drv+":\$CompactName.tag"
            Write-Host "Tagging with file $TagFileName"

            New-Item $TagFileName -type file
    
            $End = Get-Date
            $Span = New-TimeSpan -Start $Strt -End $End
            $Min = $Span.Minutes
            $Sec = $Span.Seconds
            
            if ($Discard -ne $true) {

                Write-Host "Your BMR is ready"
                Write-Host "It took $Min Min and $Sec Seconds to generate it"
                Write-Host "Next steps :"
                Write-Host "    - Remove the USB Key"
                Write-Host "    - Plug it on a $SurfaceModel"
                Write-Host "    - Boot the Surface on the USB Key"
                Write-Host ""
                Write-Host "... You should have a reimaged Surface after 20 minutes"
                return $true    
            } else {
                Write-Host "Something wrong happened"                
                Write-Host "It took $Min Min and $Sec Seconds to process"
                Write-Host "Check the logs"

                $Global:KeepExpandedDir = $true
                $Global:KeepDriversFile = $true
                $Global:KeepWimDir = $true
            }


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
                Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Discard
            }
            #Do some cleaning
            Get-Job | Remove-Job
            # Cleanup the misc Temp dir except we specify to keep them for debug purpose
            if ($Global:KeepExpandedDir -ne $true) {
                
                #Remove Expanded dir content
                write-verbose "Removing the Expanded dir"
                [String]$TMPDrvExpanded = (Join-Path -Path $env:TMP -ChildPath '\ExpandedMSIDir')
                if (Test-Path $TMPDrvExpanded) {Remove-Item $TMPDrvExpanded -Force -Recurse}
                
            }

            if ($Global:KeepDriversFile -ne $true) {
                
                #TBD Remove SDrivers files
                write-verbose "Removing the SDrivers.log and SDrivers.msi files"
                [String]$SDriversMsi = (Join-Path -Path $env:TMP -ChildPath '\SDrivers.msi')
                if (Test-Path $SDriversMsi) {Remove-Item $SDriversMsi -Force}

                [String]$SDriversLog = (Join-Path -Path $env:TMP -ChildPath '\SDrivers.log')
                if (Test-Path $SDriversLog) {Remove-Item $SDriversLog -Force}

            }

            if ($Global:KeepWimDir -ne $true) {
                
                #TBD Remove WimDir
                write-verbose "Removing the TMP Wim dir"
                [String]$TMPWimDir = (Join-Path -Path $env:TMP -ChildPath '\WimDir')
                if (Test-Path $TMPWimDir) {Remove-Item $TMPWimDir -Force -Recurse}

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
        ,
        [Alias('MakeISO')]
        [bool]$MkIso
        ,
        [Alias('TargetSKU')]
        [string]$TrgtSKU
    )
    
    begin {
        Write-Verbose "Begin procesing New-USBKey($DrvLetter,$SrcISO,$DrvRepoPath,$SurfaceModel,$TargetedOS,MakeISO=$MkIso,$TrgtSKU)"

        $Global:SkipJ1 = $False
        $Global:SkipJ2 = $False
        $Global:SkipJ3 = $False
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
            
            Set-USBKey -Drive $DrvLetter -SrcISO $SrcISO -Model $SurfaceModel -TargOS $TargetedOS -DrvRepo $DrvRepoPath -MakeISO $MkIso -Sku $TrgtSKU

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
    Helper function exported from the SurfUtil.psm1 module
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
    Helper function exported from the SurfUtil.psm1 module
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
        Write-verbose "Begin procesing Set-DriverRepo(Root=$RootRepo,Model=$SubFolder)"  
    }
  
    process {
  
        Write-Verbose "Verify $RootRepo"
        try {

            If(!(test-path $RootRepo))
            {
                Write-Host "Create $RootRepo"
                New-Item -ItemType Directory -Force -Path $RootRepo | Out-Null
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
        
        If(!(test-path $RepoPath)) {
            Write-Host "Create $RepoPath"
            New-Item -ItemType Directory -Force -Path $RepoPath | Out-Null
        }
        $RepoPath = (Get-Item -Path $RepoPath -Verbose).FullName

        Write-Host "Repo path set to $RepoPath"

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

$ModuleName = "SurfUtil"
Write-Verbose "Loading $ModuleName Module"

Export-ModuleMember -Function New-USBKey,New-IsoFile,Import-SurfaceDrivers,Import-SurfaceDB,Import-Config

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Verbose "Module $ModuleName Unloaded" }

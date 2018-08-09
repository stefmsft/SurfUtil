function IsOSVersionSupported {
    param (
        [string]$OS2Check
    )

    $returnv = $false
    $DBFileName = "$PSScriptRoot\ModelsDB.xml"
    If(test-path $DBFileName) {
        [XML]$ModelDBFile = Get-Content $DBFileName
    }
    else {
         Write-Verbose "Warning Model DB File not found"
         return $returnv
    }


    foreach ($Child in $ModelDBFile.ModelsDB.OSRelease.ChildNodes ) {

        if ($Child.ReleaseCode.tolower() -eq $OS2Check.tolower()) {
            if ($null -ne $Child.Supported) {
                if ($Child.Supported.tolower() -eq "n") {
                    return $False
                } else {
                    return $True #Supported by default
                }
            } else {
                return $True
            } #If Supported property absent then OS is Supported by default
        }
    }

}
function Get-LatestCUUrl{
    <#
    .SYNOPSIS
    Function returns the download URL as string of the latest CU for the requested Build.
    .DESCRIPTION
    Author Eric Scherlinger
        Written to enhance SurfUtil processing
    .EXAMPLE
    Get-LatestCUURL
    .EXAMPLE
    Get-LatestCUURL -TargetOS "1607"
    .PARAMETER TargetOS
    Enter the Win 10 Version number (1507,1511,1607,...) by default we use the latest version.
    .NOTES
    NAME:  Get-LatestCUURL
    AUTHOR: Eric Scherlinger
    LASTEDIT: 13/07/2018
    #>
    param
    (
        [Alias('TargetOS')]
        [string]$TargetedOS
    )

    ## Check if Target OS provided and instantiate a version filter.
    [string]$URL="Not Found"
    $versionfilter=$null
    if($TargetedOS){
        ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB
        $versionfilter= $OSReleaseHT.item($TargetedOS)
    }

    [string] $StartKB = 'https://support.microsoft.com/app/content/api/content/asset/en-us/4000816'
    ## JSON Source to all Windows Update for Win10

    ## Get the list of KBs for Win 10
    $KBs= Invoke-RestMethod -Uri $Startkb

    # Get the Latest KB either latest or based on the version filter.
    if($versionfilter){

        $LatestKB= ($kbs.links | Where-Object {$_.text -like "*$versionfilter*"})[0]
    }
    else {
        $LatestKB=$KBs.links | Where-Object {$_.id -eq $KBs.links.Count}
    }

    ## Search for the KB GUID
    $kbObj = Invoke-WebRequest -Uri "http://www.catalog.update.microsoft.com/Search.aspx?q=KB$($LatestKB.articleID)%20x64%20windows%2010"
    # Parse the Response
    $KBGUID=($kbObj.Links | Where-Object {$_.id -match "_link"}).id -replace "_link"
    $KBText=($kbObj.Links | Where-Object {$_.id -match "_link"}).innerText.ToLower()

    $i=0
    foreach ($kb in $KBGUID) {

        if ($KBText.count -eq 1) {
            $curTxt = $KBText
        } else {
            $curTxt = $KBText[$i]
        }

        if ($curTxt.Contains("cumulative")) {
            #Select only Cumulatives

            ##Creat Post Request to get the Download URL of the Update
            $Post = @{ size = 0; updateID = $kb; uidInfo = $kb } | ConvertTo-Json -Compress
            $PostBody = @{ updateIDs = "[$Post]" }

            ## Fetch and parse the download URL
            $PostRes = (Invoke-WebRequest -Uri 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody).content
            $URL= ($PostRes | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique | ForEach-Object { [PSCustomObject] @{ Source = $_.matches.value } } ).source
        }
        $i = $i + 1
    }

    ##Return the URL
    return $URL
}
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
                write-verbose "Found list of matching iso $ListMatchingISO"
                foreach ($ISOfile in $ListMatchingISO) {
                    write-verbose "Found $ISOFile"
                    $ResultISO = "$SrcISO\$ISOfile"
                }
            }

            write-verbose "Selected $ResultISO to open"

            $MountedLetter = ""
            $MountedISO = Mount-DiskImage -ImagePath $ResultISO -PassThru
            $MountedLetter = ($MountedISO | Get-Volume).DriveLetter
            $IsoMounted = $true
            Write-Verbose "ISO mounted on $MountedLetter"

            $TargetInstalWim = $MountedLetter+":\Sources\install.wim"
            Write-Verbose "Checking $TargetInstalWim"
            if (!(Test-Path $TargetInstalWim)) {
                Write-Host -ForegroundColor Red "The mounted ISO doesn't hold a valid Wim to use";
                Write-Verbose "Dismounting $ResultISO"
                DisMount-DiskImage -ImagePath $ResultISO
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
            write-host "    5 - Optimize and copy the wim to the Key"
            if ($MakeISO -eq $True) {write-host "    6 - Generate an ISO copy of your USB Key"}
            write-host ""
            write-host "Please, don't interrupt the script ...."

            $StrtAll = Get-Date
            [System.Collections.ArrayList]$JbLst = @()

######
# Phase 1 - Throw as much parallel job as we can
######

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
                        write-host "========="
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
                        write-host "========="
                        write-host ":: Step 2"
                        write-host "========="

                        [String]$SrcWimFile = (Join-Path -Path "$SrcWimLetter" -ChildPath '\sources\install.wim')
                        [String]$DstWimFile = (Join-Path -Path "$TDir" -ChildPath '\install.wim')
                        write-host "Copying $WimToRemove install.wim"
                        copy-item $SrcWimFile -Destination $DstWimFile
                        Set-ItemProperty $DstWimFile -name IsReadOnly -value $false
                        $skul = Get-WindowsImage -ImagePath $DstWimFile
                        #Get the first SKU the match the TargetedSku - Be careful on how you define sku label in the config.xml
                        $sku = ($skul | where-object {$_.imagename.tolower() -like $TargetSKU.tolower()})[0]
                        $indx = $sku.imageindex
                        Write-verbose "Found [$TargetSKU] at index $indx"

                        if ($null -eq $sku) {

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
                        write-host "========="
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

######
# Job 4 Gather the latest CU for the targeted Windows Version
######

            if ($Global:SkipJ4 -ne $true) {

                $localp = (Get-Item -Path ".\" -Verbose).FullName
                $mp = "$localp\SurfUtil.psm1"

                $localp = (Get-Item -Path ".\" -Verbose).FullName
                $LCUPathDir = "$localp\WindowsCU"

                Write-Verbose "Firing Up Job 4 ..."
                $JobGetCU = Start-Job -Name "Job4" -ScriptBlock {

                    param ([String] $os, [String] $LocalCUPathDir, [String] $IDMPath, [String] $v)

                        $VerbosePreference=$v
                        write-verbose "Call Job 4 ($os,$LocalCUPathDir,$IDMPath)"
                        $Strt = Get-Date
                        write-host "========="
                        write-host ":: Step 4"
                        write-host "========="

                        $ret = import-module $IDMPath

                        Write-Verbose "Calling : Get-LatestCU -WindowsVersion $os -LocalCUDir $LocalCUPathDir"
                        $ret = Get-LatestCU -WindowsVersion $os -LocalCUDir $LocalCUPathDir

                        if ($null -eq $ret) {
                            write-host "Step 4 : Operation Failed"
                            write-host "         Gathering CU failed"
                        }

                        $End = Get-Date
                        $Span = New-TimeSpan -Start $Strt -End $End
                        $Min = $Span.Minutes
                        $Sec = $Span.Seconds
                        Write-Host "CU gathered in $Min Min and $Sec Seconds"
                        write-host "============="
                        write-host ":: End Step 4"
                        write-host "============="

                        return $ret

                } -ArgumentList $TargetedOS ,$LCUPathDir , $mp, $VerbosePreference

            } else {

                Write-verbose "Job 4 Skipped"
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

######
# Phase 2 - Start some synchrone work (Wim manipulation)
######
            $StatusInfo = "Success"

            if ($WimIndx -ge 0) {

                ## Preliminary test on the ExpandedMSIDir to avoid the useless mount of the wim after
                #Bug Fix - Corp changed the structure of the expanded dir
                #          No more PlateFormInstaller directory but a SurfaceUpdate directory
                #          To avaoid futur break I'm just checking that the dir has contents
                [String]$DrvExpandRoot = (Join-Path -Path $env:TMP -ChildPath '\ExpandedMSIDir')

                $ExpDirCnt = (Get-ChildItem -Path $DrvExpandRoot).count
                Write-verbose "Found $ExpDirCnt item in the expanded MSI root directory"

                If($ExpDirCnt -gt 0) {
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
                    $ReadMe = (Get-ChildItem -Recurse -Path $DrvExpandRoot -filter ReadMe.txt).FullName

                    if (-not ([string]::IsNullOrEmpty($ReadMe))) {

                        Copy-Item -Path $ReadMe -Destination ($Drv+":\ReadMe.txt")
                    }

                    ## Inject the latest CU
                    if ($Global:SkipJ4 -ne $true) {
                        Wait-Job $JobGetCU | Out-Null
                        $CUFile = Receive-Job $JobGetCU
                        Write-Verbose "CU File is $CUFile"

                        if (Test-Path $CUFile) {

                            Add-WindowsPackage -Path $MntDir -PackagePath $CUFile -PreventPending -NoRestart -IgnoreCheck -LogPath $LogPath | Out-Null
                            Write-Host "Latest Cumulative Update injected in the new Wim"

                        } else {

                            Write-Host "No Cumulative file found - Nothing to apply"

                        }
                    }

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
                    $StatusInfo = "Job 3 did not correctly expand the drivers in [$DrvExpandRoot]"

                }

                $WimMounted = $false

                if ($Discard -eq $true) {

                    $CompactName = "ERROR"
                    if ($StatusInfo -eq "Success") {

                        $StatusInfo = "Something went wrong with the Drivers Injection or Wim mounting"
                    }
                } else {

                    $CompactName = "BMR-" + $SurfaceModel.replace(" ","").toupper() + $TargetedOS
                    #Export the image to a more compact one
                    [String]$ExptWimFile = (Join-Path -Path $TmpTDir -ChildPath 'install_serviced.wim')
                    $Ret = Export-WindowsImage -SourceImagePath $MntWimFile -SourceIndex $WimIndx -DestinationImagePath $ExptWimFile -DestinationName $CompactName -ScratchDirectory $ScrtchDir -LogPath $LogPath
                    [String]$FinalSwmFile = (Join-Path -Path ($Drv + ":") -ChildPath '\sources\install.swm')
                    $ret = Split-WindowsImage -FileSize 3500 -ImagePath $ExptWimFile -SplitImagePath $FinalSwmFile -CheckIntegrity -ScratchDirectory $ScrtchDir
                }
            } else {

                $CompactName = "ERROR"
                $StatusInfo = "Job 2 did not find a valid index for [$TSku]"
            }

            Write-Verbose "Dismounting $ResultISO"
            DisMount-DiskImage -ImagePath $ResultISO
            $IsoMounted = $false

            # Generate ISO file
            if (($MakeISO -eq $true) -and ($Discard -ne $true)) {
                $TCIsoName = ".\" + $CompactName + ".iso"
                Get-ChildItem ($Drv + ":\") | New-IsoFile -Media 'DVDPLUSR_DUALLAYER' -Title $CompactName -Path $TCIsoName -Force

            }

            #Write tag file on completion
            $TagFileName = $Drv+":\$CompactName.tag"
            Write-Host "Tagging with file $TagFileName"

            New-Item $TagFileName -type file -value $StatusInfo
            $End = Get-Date
            $Span = New-TimeSpan -Start $StrtAll -End $End
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
                Write-Host "Check the logs in $env:TMP\ExpandedMSIDir"
                write-host "            or in $env:TMP\Wimdir\DISM.log"
                write-Host "Or run the script with -Log $True"

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
                DisMount-DiskImage -ImagePath $ResultISO
            }
            # Check if the Wim need to be dismounted
            if ($WimMounted -eq $true) {
                Write-Verbose "Dismounting Wim after Exception"
                Dismount-WindowsImage -Path $MntDir -ScratchDirectory $ScrtchDir -LogPath $LogPath -Discard
            }
            #Do some cleaning
            Get-Job | Receive-Job
            Remove-Job *
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
function Update-USBKey {
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
        Write-Verbose "Begin procesing Update-USBKey($DrvLetter,$SrcISO,$DrvRepoPath,$SurfaceModel,$TargetedOS,MakeISO=$MkIso,$TrgtSKU)"
    }

    process {

        try {

            #Do something

        }
        catch [System.Exception] {
            Write-Host -ForegroundColor Red $_.Exception.Message;
            return $False
        }

    }

    end {
        Write-Verbose "End procesing Update-USBKey"
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
        ,
        [Alias('Log')]
        [string]$LogAsk
    )

    begin {
        Write-Verbose "Begin procesing New-USBKey($DrvLetter,$SrcISO,$DrvRepoPath,$SurfaceModel,$TargetedOS,MakeISO=$MkIso,$TrgtSKU,Log=$LogAsk)"

        $Global:SkipJ1 = $False
        $Global:SkipJ2 = $False
        $Global:SkipJ3 = $False
        $Global:SkipJ4 = $False

        $Global:KeepExpandedDir = $false
        $Global:KeepDriversFile = $false
        $Global:KeepWimDir = $false
    }

    process {

        try {

            #Set logging to verbose if asked
            if ($LogAsk -eq $True) {

                $VerbosePreference = "Continue"
                $DebugPreference = "Continue"

            }

            #Display in verbose the Skip Jobs flags
            Write-Verbose "Skip Job 1 : $Global:SkipJ1"
            Write-Verbose "Skip Job 2 : $Global:SkipJ2"
            Write-Verbose "Skip Job 3 : $Global:SkipJ3"
            Write-Verbose "Skip Job 4 : $Global:SkipJ4"

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
         ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB
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
        ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB
    }

    process {

        try {

            [System.Collections.ArrayList]$CurLst = @()

            if ($DrvModel -ne "") {
                $urldrv = $SurfModelHT[$DrvModel.tolower()]
                $SearchPat = $SurfModelPS[$DrvModel.tolower()]
                if ($null -eq $urldrv) {
                    Write-Host -ForegroundColor Red "Unknown Surface Model for the script : [$DrvModel]"
                    return $false
                }
            }

            if (($null -ne $OVers) -and ($OVers -ne "")) {
                $InternalR = $OSReleaseHT[$OVers.tolower()]
                if ($null -eq $InternalR) {
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
                if ($null -ne $href) {
                    if ($href -like $SearchPat ) {
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
        ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB

    #>
        $DBFileName = "$PSScriptRoot\ModelsDB.xml"
        If(test-path $DBFileName) {
            [XML]$ModelDBFile = Get-Content $DBFileName
        }
        else {
             Write-Verbose "Warning Model DB File not found"

        }

        $ModelsHT = @{}   # empty models hashtable for Drv Url
        $ModelsHTSP = @{}   # empty models hashtable for Drv PAttern filter

        foreach ($Child in $ModelDBFile.ModelsDB.SurfacesModels.ChildNodes ) {

                $ModelsHT.Add($Child.ID.tolower(),$Child.Drivers.url.tolower())
                if ($null -eq $Child.Drivers.searchpattern) {
                    $ModelsHTSP.Add($Child.ID.tolower(),"*Win10*.msi")
                } else {
                    $ModelsHTSP.Add($Child.ID.tolower(),$Child.Drivers.searchpattern)
                }
        }

        $OSHT = @{}   # empty models hashtable

        foreach ($Child in $ModelDBFile.ModelsDB.OSRelease.ChildNodes ) {

                $OSHT.Add($Child.ReleaseCode.tolower(),$Child.InternalCode.tolower())
        }

    $CtxData = ($ModelsHT,$OSHT,$ModelsHTSP)

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
            write-verbose $FulltargetFile

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
        ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB
        $DefaultFromConfigFile = Import-Config
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
            Write-Debug "Sta tus of Get-RemoteDriversInfo $status"
            $status = Get-LocalDriversInfo -RootRepo $RepoPath -DrvModel $Model -OSTarget $WindowsVersion
            Write-Debug "Sta tus of Get-LocalDriversInfo $status"
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
function Get-LatestCU {
    <#
    .SYNOPSIS

    .DESCRIPTION
    Get the latest CU for a giving Windows 10 version

    .EXAMPLE
    Get-LatestCU -WindowsVersion 1803

    .PARAMETER WindowsVersion
    Targeted Version of Windows 10
    Supported value are listed in the ModelsDB.XML file.

    #>
    [CmdletBinding()]
    param
    (
        [string]$WindowsVersion,
        [string]$LocalCUDir=".\WindowsCU"
    )

    $CUUrl = (Get-LatestCUUrl($WindowsVersion))
    Write-Verbose $CUUrl

    # Verify if the Local dir for CU is present
    write-verbose "Testing if $LocalCUDir exists"
    If(!(test-path $LocalCUDir)) {

        write-verbose "Create $LocalCUDir directory"
        New-Item -ItemType Directory -Force -Path $LocalCUDir | out-null

    }

    #Create the local name for the CSU
    #Scheme : Window-XXXX-KBXXXXXXXXXX-X64.msu

    $sp1 = $CUUrl -split '-'
    $KBName = $sp1[1]
    $Filename = "Windows-$WindowsVersion-$KBName-X64.msu"
    Write-Verbose "Local filename will be : $Filename"

    $FullName = "$LocalCUDir\$Filename"

    If(!(test-path $FullName)) {

        #download the Msu file
        $Strt = Get-Date
        Write-Host "Start Downloading latest CU .................... "

        Get-MSIFile -Link $CUUrl -LPath $LocalCUDir -File $Filename
        $End = Get-Date
        $Span = New-TimeSpan -Start $Strt -End $End
        $Min = $Span.Minutes
        $Sec = $Span.Seconds

        Write-Host "Downloaded in $Min Min and $Sec Seconds"
    } else {

        Write-Host "Latest CU for $WindowsVersion already present locally"
    }


    return $FullName

}

$ModuleName = "SurfUtil"
Write-Verbose "Loading $ModuleName Module"

Export-ModuleMember -Function New-USBKey,New-IsoFile,Import-SurfaceDrivers,Import-SurfaceDB,Import-Config,Get-LatestCU

$ThisModule = $MyInvocation.MyCommand.ScriptBlock.Module
$ThisModule.OnRemove = { Write-Verbose "Module $ModuleName Unloaded" }

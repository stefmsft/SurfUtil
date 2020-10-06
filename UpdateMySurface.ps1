Param(  [string]$ForceModel,
        [switch]$Log,
        [switch]$CheckOnly,
        [switch]$Yes        
    )


    function Get-EstimatedModel {
        <#
        .SYNOPSIS
        Short description
        
        .NOTES
        
        Based on
        Ref : https://docs.microsoft.com/en-us/surface/surface-system-sku-reference
        
        Device	                    System Model	    System SKU
        ----------------------------------------------------------
        Surface 3 WiFI	            Surface 3	        Surface_3
        Surface 3 LTE AT&T	        Surface 3	        Surface_3_US1
        Surface 3 LTE Verizon	    Surface 3	        Surface_3_US2
        Surface 3 LTE North America	Surface 3	        Surface_3_NAG
        
        Surface Pro	                Surface Pro	        Surface_Pro_1796
        Surface Pro with LTE    	Surface Pro	        Surface_Pro_1807
        Surface Pro 6 Consumer	    Surface Pro 6	    Surface_Pro_6_1796_Consumer
        Surface Pro 6 Commercial	Surface Pro 6	    Surface_Pro_6_1796_Commercial
        Surface Pro 7	            Surface Pro 7	    Surface_Pro_7_1866
        Surface Pro X	            Surface Pro X	    Surface_Pro_X_1876
        
        Surface Book 2 13"	        Surface Book 2	    Surface_Book_1832
        Surface Book 2 15"	        Surface Book 2	    Surface_Book_1793
        Surface Book 3 13"	        Surface Book 3	    Surface_Book_3_1900
        Surface Book 3 15"	        Surface Book 3	    Surface_Book_3_1899
        
        Surface Go LTE Commercial	System Go	        Surface_Go_1825_Commercial
        Surface Go Consumer	        Surface Go	        Surface_Go_1824_Consumer
        Surface Go Commercial	    Surface Go	        Surface_Go_1824_Commercial
        Surface Go 2	            Surface Go 2	    Surface_Go_2_1927
        
        Surface Laptop	            Surface Laptop	    Surface_Laptop
        Surface Laptop 2 Consumer	Surface Laptop 2	Surface_Laptop_2_1769_Consumer
        Surface Laptop 2 Commercial	Surface Laptop 2	Surface_Laptop_2_1769_Commercial
        Surface Laptop 3 13" Intel	Surface Laptop 3	Surface_Laptop_3_1867:1868
        Surface Laptop 3 15" Intel	Surface Laptop 3	Surface_Laptop_3_1872
        Surface Laptop 3 15" AMD	Surface Laptop 3	Surface_Laptop_3_1873
        #>
        param
        (
            [string]$SKUModel
        )
    
        if ($SKUModel -eq "Surface_Pro_1807") {
            $SurfaceModel = "Surface Pro LTE"
        }
        elseif ($SKUModel -eq "Surface_Pro_1796") {
            $SurfaceModel = "Surface Pro"
        }
        elseif ($SKUModel -eq "Surface_Pro_6_1796_Consumer") {
            $SurfaceModel = "Surface Pro 6"
        }
        elseif ($SKUModel -eq "Surface_Pro_6_1796_Commercial") {
            $SurfaceModel = "Surface Pro 6"
        }
        elseif ($SKUModel -eq "Surface_Pro_7_1866") {
            $SurfaceModel = "Surface Pro 7"
        }
        elseif ($SKUModel -eq "Surface_Pro_X_1876") {
            Write-Host -ForegroundColor Red "No MSI available for Surface Pro X - Use Microsoft Update instead"
            $SurfaceModel = "NA"
        }
        elseif ($SKUModel -eq "Surface_Go_1825_Commercial") {
            $SurfaceModel = "Surface Go LTE"
        }
        elseif ($SKUModel -eq "Surface_Go_1824_Consumer") {
            $SurfaceModel = "Surface Go"
        }
        elseif ($SKUModel -eq "Surface_Go_1824_Commercial") {
            $SurfaceModel = "Surface Go"
        }
        elseif ($SKUModel -eq "Surface_Go_2_1927") {
            $SurfaceModel = "Surface Go 2"
        }
        elseif ($SKUModel -eq "Surface_Book_1832") {
            $SurfaceModel = "Surface Book 2"
        }
        elseif ($MachiSKUModelneModel -eq "Surface_Book_1793") {
            $SurfaceModel = "Surface Book 2"
        }   
        elseif ($SKUModel -eq "Surface_Book_1832") {
            $SurfaceModel = "Surface Book 2"
        }
        elseif ($SKUModel -eq "Surface_Book_3_1900") {
            $SurfaceModel = "Surface Book 3"
        }
        elseif ($SKUModel -eq "Surface_Book_3_1899") {
            $SurfaceModel = "Surface Book 3"
        }
        elseif ($SKUModel -eq "Surface_Pro_4") {
            $SurfaceModel = "Surface Pro 4"
        }
        elseif ($SKUModel -eq "Surface_Pro_3") {
            $SurfaceModel = "Surface Pro 3"
        }
        elseif ($SKUModel -eq "Surface_Studio") {
            $SurfaceModel = "Surface Studio"
        }
        elseif ($SKUModel -eq "Surface_Laptop") {
            $SurfaceModel = "Surface Laptop"
        }
        elseif ($SKUModel -eq "Surface_Laptop_2_1769_Consumer") {
            $SurfaceModel = "Surface Laptop 2"
        }
        elseif ($SKUModel -eq "Surface_Laptop_2_1769_Commercial") {
            $SurfaceModel = "Surface Laptop 2"
        }
        elseif ($SKUModel -eq "Surface_Laptop_3_1867:1868") {
            $SurfaceModel = "Surface Laptop 3 INTEL"
        }
        elseif ($SKUModel -eq "Surface_Laptop_3_1872") {
            $SurfaceModel = "Surface Laptop 3 INTEL"
        }
        elseif ($SKUModel -eq "Surface_Laptop_3_1873") {
            $SurfaceModel = "Surface Laptop 3 AMD"
        }
        else {
            #Fallback
            $SurfaceModel = (Get-WmiObject -namespace root\wmi -class MS_SystemInformation | select-object SystemProductName).SystemProductName
        }
    
        return $SurfaceModel
    }

    Import-Module "$PSScriptRoot\SurfUtil.psm1" -force | Out-Null

try {

    #Verifiy if ran in Admin
    $IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    if ($IsAdmin -eq $False) {
        Write-Host -ForegroundColor Red "Please use this script in an elevated Admin context"
        return $false
    }

    if ($Log -eq $true) {

        $OldVerboseLevel = $VerbosePreference
        $OldDebugLevel = $DebugPreference

        $VerbosePreference = "Continue"
        $DebugPreference = "Continue"

    }

    ($SurfModelHT,$OSReleaseHT,$SurfModelPS) = Import-SurfaceDB

    if ($ForceModel -eq "") {   

        $MachineModel = (Get-WmiObject -namespace root\wmi -class MS_SystemInformation | select-object SystemSKU).SystemSKU
        $SurfaceModel = Get-EstimatedModel ($MachineModel)

    } else {

        $MachineModel = $ForceModel
        $SurfaceModel = Get-EstimatedModel ($MachineModel)
    }

    if ($SurfaceModel -eq "") {
        Write-Host -ForegroundColor Red "Surface Model unidentified";
        return $False    
    }
    elseif ($SurfaceModel -eq "NA") {
        return $False
    }
    else {
        Write-Verbose "Model : $SurfaceModel"
    }

    $DefaultFromConfigFile = Import-Config

    $localp = (Get-Item -Path ".\" -Verbose).FullName

    $RepoPath = $DefaultFromConfigFile["RootRepo"]
    if ($null -eq $RepoPath) {
        $RepoPath = "$localp\Repo"
    }

    If(!(test-path $RepoPath)) {
            New-Item -ItemType Directory -Force -Path $repopath | out-null
    }

    $LocalRepoPathDir = resolve-path $RepoPath
    write-debug "Full Path Repo is : $LocalRepoPathDir"

    $ios = (Get-WmiObject Win32_OperatingSystem | select-object BuildNumber).BuildNumber
    $os = (($OSReleaseHT.GetEnumerator()) | Where-Object { $_.Value -eq $ios }).Name

    if ($CheckOnly) {
        $Apply = $False
        Write-Verbose "The drivers won't be applied to the local machine"
    } else {
        $Apply = $True
        Write-Verbose "The drivers will be applied to the local machine after download"
    }

    if ($Yes -eq $false) {

        $message  = "Ready to download of the latest drivers for $SurfaceModel"
        $question = 'Are you sure you want to proceed?'

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

        $decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
        if ($decision -eq 1) {

            Write-Host 'Action canceled ...'
            return $False

        }

    }

    $ret = Import-SurfaceDrivers -Model $SurfaceModel -WindowsVersion $os -RepoPath $LocalRepoPathDir -Apply $Apply
    if ($ret -eq $False) {
        Write-Host "No Drivers found for the current OS ... Looking for previous versions"
        $ret = Import-SurfaceDrivers -Model $SurfaceModel -RepoPath $LocalRepoPathDir -Apply $Apply
    }

}
catch [System.Exception] {
    Write-Host -ForegroundColor Red $_.Exception.Message;
    return $False
}
finally {
    if ($Log -eq $true) {

        write-verbose "Re establish initial verbosity"
        $VerbosePreference = $OldVerboseLevel
        $DebugPreference = $OldDebugLevel

    }
}

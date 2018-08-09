# Introduction

**SurfUtil** or *Surface Utilities* is a set of tools that eases the management of Surface Tablets.
In its first version the focus is set on the generation of bootable installation USB keys specific to each Surface model.

# Description

The core part is held inside of a module (SurfUtil.psm1)
Then you'll find a set of little script functions dedicated to spécific actions. This is those function that most poeple will use and I'm going to detail then below.

## Drivers related functions

  The goal here is to check Microsoft donwload driver web site ( [http://aka.ms/surfacedrivers](http://aka.ms/surfacedrivers) ) for the latest version of a driver set and gather it to a local repository if missing.
  A driver set is a msi file regularly pushed online containing the cumulative updated drivers and firmware for a surface model.

  By default the local repository directory is named ".\Repo" and have the following structure :
+ .\Repo
  + Surface model
    + Targeted OS Version

Two functions are available in this field:

+ UdateRepo.ps1
It gathers all the Surface Models listed in the ModelsDB.XML file and verifies if we have locally (under .\Repo by default) the latest version of the Surface driver set available online.
>_Note, that for some Surface Model, you can have multiple "latest driver set" because some are published explicitly for a minimun version of the Targeted OS._

+ UdateMySurface.ps1
When ran on a Surface Tablet, it gathers the device info (model, OS version, ... ) so it can verifies if we have latest drivers set locally and if not, it gather it and apply it to your Surface.
This is equivalent of forcing a download from WU or checking online if a new driver set msi is available to apply.
You can do this in a one line of powershell command and even can think of creating a scheduled job to make regular check on your behalf.

## Bare Metal Recovery Master functions

  The goal of those functions (only one for now but more will come) is to generate an bootable USB key to achieve actions like reinstalling a slipstreamed version (OS+Driver Set) targeted to a specific version of Windows 10 (see below)
  + MakeBMR.ps1
As described above, this will allow you the generate an bootable USB Key for a model of Surface. This key will allow  you to reinstal the machine in arround 20 min.

Two important prerequisit
- You need to provide the letter of the USB key you inserted in your machine.
- The function expect to find an ISO holding the targeted version of the OS in the .\iso subdirectory.
  - The script use the string __*"\*windows_10\*_TargetedOSVersion\*"*__ as a search filter.
    - For instance if my **.\iso** directory holds the file *en_windows_10_business_editions_version_1803_updated_march_2018_x64_dvd_12063333.iso* freshly download from msdn for instance, the script will be happy to generate USB keys with 1803 as a targeted OS Version.
    >Note also that the iso have to held at least one *Windows 10 Pro Sku* distribution. This restriction might change in the future.

So for example, if you wish to generate an bootable Bare Metal Recovery key based on Windows 10 1803 for your Surface Pro 3 you just have to use the following command :

      .\MakeBMR -Drive D -SurfaceModel "Surface Pro3" -WindowsVersion 1803

+ Where D is the letter holding my 8Go USB Key
+ You can add also the *-MkISO $true* parameter to generate an ISO copy of your key for future reuse.

Keep in mind that there is no magic in the tool, It just gathers the required OS file from the iso then gathers the latest driver from the net finaly it mixes that and optimizes the resulting media size.  It also do the spliting of wim files when needed. The result is an optimized slipstreamed version of a Windows installation key dedicated to the Surface Model requested.

    Here is the output of the command expected

        Create a BMR Key for [Surface Pro3] / [windows 10 pro 1803]

        If you agree, the external drive labeled [BMR 1803] will be formated
        Are you sure you want to proceed?
        [Y] Yes  [N] No  [?] Aide (la valeur par défaut est « N ») : y
        The next steps are :
            1 - Format the target and copy windows files
            2 - Prepare the Wim File
            3 - Inject Surface Drivers
            5 - Optimize and copy the wim to the Key
            6 - Generate an ISO copy of your USB Key

        Please, don't interrupt the script ....
        Mounting Image...
        Drivers injected in the new Wim
        Committing Changes and Dismounting Image...
        Your BMR is ready

        It took 36 Min and 8 Seconds to generate it
        Next steps :
            - Remove the USB Key
            - Plug it on a Surface Pro3
            - Boot the Surface on the USB Key

        ... You should have a reimaged Surface after 20 minutes

## Supported model
Here is below the list of Surface Model supported with the current version of the tool :
+ 'Surface Pro'
+ 'Surface Pro LTE'
+ 'Surface Book'
+ 'Surface Book 2 [13¨/15¨]'
+ 'Surface Studio' (*Still under validation process for MakeBMR - Be careful)
+ 'Surface LapTop'
+ 'Surface Pro4'
+ 'Surface Pro3'

## Supported Windows 10 OS Versions
Basically the supported OS are the OS version for witch a supported driver package (msi file) is available. The qualifying OS Version are also reflecting the current supported version of Windows 10.
A the date of the latest release this is :
+ TH1 (1507)
+ TH2 (1511)
+ RS1 (1607)
+ RS2 (1703)
+ RS3 (1709)
+ RS4 (1803)

As for the Drivers part, the available drivers version online will drive what the tool support ... For the boot key generation, you are in charge of providing the OS distribution you wish to use (iso file).The result will be as good as the quality of the ingredient you use with the tool.

# Getting Started

You can acces to the recent release of the SurfUtil set of scritp here :

https://github.com/stefmsft/SurfUtil/releases

You can also get in sync with the latest updates thru Git by cloning the following url :

https://github.com/stefmsft/SurfUtil.git

1.	Installation process

    There's no installation process per say. The available scripts loads the required module before their execution.

    *A script named LoadAndCheck.ps1 allows to load SurfUtil helper modules without executing any actions.

    In this file you can change the value of the **$VerbosePreference** variable to control the level of verborsity you wich to have during the execution of the tool functions

        If set to "Continue" it will show verbose info

        Set it back to "SilentlyContinue" to remove the verbosity

    After modification of the file, do a run of :

        . .\LoadAndCheck.ps1


2.	Software dependencies

    * Windows 10
    * Powershell V5

3.	Latest releases

    Release 1.1

    This release include :
    - A new *-Log $True* parameter to MakeBMR.ps1 to ease debugging problems
    - MakeBMR.ps1 Automatically gather the latest Cumulative Update of the OS in the WIM
    - A search pattern attribute (*searchpattern=*) is added in the MadelDB.xml file for the Drivers tag

    Release 1.0

    This first release provide :
    - The first 3 ps1 script described above (MakeBMR,UpdateMySurface,UpdateRepo)

4.	Artefacts
    * SurfUtil.psm1
    * UpdateMySurface.ps1
    * UpdateRepo.ps1
    * MakeBMR.ps1
    * config.xml
    * ModelsDB.xml

# Build and Test
Unitest are conducted by pester. They are run online thru the build process thu the tools building process (VSTS). This is part of a Continuous integration process. The release process is trigger manually then and it mainly consist to a synchronisation of the code tree to GitHub for public availability.

To get the latest version of Pester use the following command :

>*Install-Module -Name Pester -Force -SkipPublisherCheck*

Part of the released code you'll see file with .test.ps1 extensions. Those are the unitests for the project.
To manually trigger the pester tests you can run :

>*Invoke-Pester *nameofthescript*.test.ps1

You can use wildcard (*.test.ps1) to run all the test.

# Contribute
All the Microsoft Surface TSP/GBB are welcome to contribute directly to the internal VSTS project. Send me a mail so I can add you to the repo.

External contributor are welcome to submit me pull requests as well.

# Known Issue

- Download time is slow due to the API used for asynchronous method (progress bar) - This will be addressed after feature completion (Work Item : 1)
- The download logic need to be fix when -Apply $True parameter is used. Today it downloads all Target Version OS before applying the latest.

# History and thanks
This project started with 2 influences

1. The script build by a peer colleague (__Casey Hill__). I've used then for a while and the indeniable value there is the part that allow to guess the URL for Surface firmware/driver update online. When I decided to go a step further on what was available, I learned a lot from Casey scripts for the part.

2. All the work done by the team develloping and maintening the PDT project. PDT is a set of Powershell Tool that allow anyone to generate a System Center lab environment. I took a lot of their inspiring work to speed mine on generating automated booting media.
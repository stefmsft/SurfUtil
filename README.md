# Introduction 
The project is about providing a set of tools that allows the generation of bootable installation key specific to Surface.

# Description
The project is segmented in 2 phases
- Phase One

  The production of the right tooling to gather, expand, apply Surfaces Drivers.
  
  The goal is to check the donwload driver web site ( [http://aka.ms/surfacedrivers](http://aka.ms/surfacedrivers) ) for more up to date version of a driver set and gather it to a local repository of reference.
  
  By default the Repository directory is named .\Repo and have the following structure :
    * .\Repo
      * *Surface model*
        * *Targeted OS Version*
- Phase Two
  
  TBD

## Supported model
Here is below the list of Surface Model supported with the current version of the tool :
+ 'Surface Pro'
+ 'Surface Pro LTE'
+ 'Surface Book'
+ 'Surface Book 2 [13¨/15¨]'
+ 'Surface Studio'
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

As for the Drivers part, the available drivers version online will drive what the tool support ... For the boot key generation, you are in charge to provide the OS distribution you wish to use.The result will be as good as the quality of the ingredient you use with the tool.

# Getting Started
You can acces to the recent release of the BMRGen set of scritp here :

https://github.com/stefmsft/BMRGen/releases

You can also get in sync with the latest updates thru Git by cloning the following url :

https://github.com/stefmsft/BMRGen.git

1.	Installation process

    There's no instal process per say. The available scripts load the required module before their execution.
    
    *A script named LoadAndCheck.ps allows to load BMRGen helper modules without executing any actions.*

2.	Software dependencies

    * Windows 10
    * Powershell V5

3.	Latest releases
    0.3

4.	Artefacts
    1. Phase One
        - Import-SurfaceDrivers.psm1
        - UpdateMySurface.ps1
        - UpdateRepo.ps1
    2. Phase Two
        - TBD

# Build and Test
Unitest are conducted by pester. They are run online thru the build process thu the tools building process (VSTS). This is part of a Continuous integration process. The release process is trigger manually then and it mainly consist to a synchronisation of the code tree to GitHub for public availability.

To get the latest version of Pester use the following command :

>*Install-Module -Name Pester -Force -SkipPublisherCheck*

Part of the released code you'll see file with .test.ps1 extensions. Those are the uni tests for the project.

# Contribute
All the Microsoft Surface TSP/GBB are welcome to contribute directly to the internal VSTS project. Send me a mail so I can add you to the repo.

External contributor are welcome to submit me pull requests as well. 

# History and thanks
This project started with 2 influences

1. The script build by a peer colleague (__Casey Hill__). I've used then for a while and the indeniable value there is the part that allow to guess the URL for Surface firmware/driver update online. When I decided to go a step further on what was available, I learned a lot from Casey scripts for the part.

2. All the work done by the team develloping and maintening the PDT project. PDT is a set of Powsrshell Tool that allow anyone to generate a System Center lab environment. I took a lot of their inspiring work to speed mine on generating automated booting media. 
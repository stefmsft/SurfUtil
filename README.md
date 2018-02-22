# Introduction 
The project is about gatheringa set of tools that allows to generate easily a bootable key to install from scratch a Surface

# Getting Started
To use the tools first clone the git content
1.	Installation process
    git clone https://stephsau.visualstudio.com/_git/BMRGen
    cd BMRGen
    setup.ps1
    use the cmdlets, like Gen-BMR

2.	Software dependencies
    The creation of the key depends on the avbailabilty of the targeted windows WIM on the machine running the tool
    Powershell V5

3.	Latest releases
    0.1
4.	API references
    Prepare-BMR
    Gen-BMR

# Build and Test
All the code is written in Powershell, so there's no build process
So unitest will be provided later

# Contribute
All the Microsoft Surface TSP/GBB are welcome to contribute. Send me a mail so I can add you to the repo. Use pull request to submit your work. 

# History and thanks
This project started with 2 influences
1. The script build by a peer colleague (Casey Hill). I've used then for a while and the indeniable value there is the part that allow to guess the URL for Surface firmware/driver update online. When I decided to go a step further on what was available, I learned a lot from Casey scripts for the part.

2. All the work done by the team develloping and maintening the PDT project. PDT is a set of Powsrshell Tool that allow anyone to generate a System Center lab environment. I took a lot of their inspiring work to speed mine on generating automated booting media. 
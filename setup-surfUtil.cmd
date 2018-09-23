Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
choco upgrade chocolatey
choco install git -y -x
mkdir SurfUtil
"C:\Program Files\Git\cmd\git.exe" clone https://github.com/stefmsft/SurfUtil.git .\SurfUtil\

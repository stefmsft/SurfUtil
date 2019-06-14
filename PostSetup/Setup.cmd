REM To fix https://support.microsoft.com/en-us/help/4090656/issue-deploy-custom-image-or-new-windows-to-surface-without-vc-redist
REM Download vc_redist.x64.exe from https://aka.ms/vs/15/release/vc_redist.x64.exe
REM uncomment next line
REM vc_redist.x64.exe /install /quiet /norestart

REM To install Surface Dock Updater
REM Download https://download.microsoft.com/download/8/2/E/82EEFB07-1AB3-4557-B654-B34D64C9DD94/Surface_Dock_Updater_v2.22.139.0.msi
REM uncomment next line
MSIEXEC /norestart /qn /log C:\SurfaceDrivers\SurfaceDockUpdater.log /i Surface_Dock_Updater_v2.22.139.0.msi
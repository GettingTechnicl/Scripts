@echo off

:REM check if "HKLM\Software\Sentinel Labs" registry key is present
reg query "HKLM\Software\Sentinel Labs"

:REM if "HKLM\Software\Sentinel Labs" registry key is present, it means that sentinel has already been installed on this host, so go to the INSTALLED switch of the script
IF %ERRORLEVEL% == 0 goto INSTALLED

:REM Copy SentinelInstaller_windows.msi installer from SYSVOL share to local TEMP folder WORKSTATION
copy \\std\sysvol\std.local\scripts\SentinelOne\SentinelInstaller_windows.msi c:\windows\temp\ /Z /Y

:REM install msi package
msiexec /i "c:\windows\temp\SentinelInstaller_windows.msi" /q /norestart SITE_TOKEN="ps3GpmsPqogCBKF0ANnRhmUVptppZlKPMncnl2CGNG6cbaHia3yRHw6aWRb12AeDSj5NpabG1T4A6XPWzOsHt62jAgwK8IL5l0JibeWa"

:REM if install is ok go to OK switch
IF %ERRORLEVEL% == 0 goto OK

:REM if install fails go to ERROR switch
goto ERROR

:INSTALLED
echo "Already Installed"
goto END

:ERROR
echo "Install Error"
goto END

:OK
echo "Install OK"

:END
## CPEF

#Download AW Agent

$awAgent = $env:TEMP + "\arcticwolfagent-2022-03_52.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CRckgL9drec48ei/download/arcticwolfagent-2022-03_52.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awAgent)


#Download Customer.json

$awJson = $env:TEMP + "\customer.json"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/gbX2TwtQFRXkTZz/download/customer.json"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awJson)


#Download Sysmon.exe

$awSysmon = $env:TEMP + "\Sysmon.exe"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/MEJne7CnypQt2R7/download/Sysmon.exe"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmon)


#Download Sysmon64.exe

$awSysmon64 = $env:TEMP + "\Sysmon64.exe"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/aQazK2Z6mngXRwc/download/Sysmon64.exe"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmon64)


#Download AW SysmonAssistant

$awSysmonAssistant = $env:TEMP + "\sysmonassistant-1_0_1.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/9aZQiyYk8KHag9P/download/sysmonassistant-1_0_1.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmonAssistant)


# Install Sysmon via AW Assistant
Start-Process -FilePath $Msiexec -ArgumentList "/i $awSysmonAssistant","/Q","/NORESTART" -Wait 

# Install AW Agent
Start-Process -FilePath $Msiexec -ArgumentList "/i $awAgent","/Q","/NORESTART" -Wait 
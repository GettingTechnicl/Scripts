#Download AW Agent

$awAgent = $env:TEMP + "\arcticwolfagent-2022-03_52.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CHBMJw4FAqfdtQ7/download?path=%2F&files=arcticwolfagent-2022-03_52.msi&downloadStartSecret=pwgei3ybxmf"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awAgent)


#Download Customer.json

$awJson = $env:TEMP + "\customer.json"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CHBMJw4FAqfdtQ7/download?path=%2F&files=customer.json&downloadStartSecret=ophnuc5ls7"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awJson)


#Download Sysmon.exe

$awSysmon = $env:TEMP + "\Sysmon.exe"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CHBMJw4FAqfdtQ7/download?path=%2F&files=Sysmon.exe&downloadStartSecret=mrh54cpmseh"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmon)


#Download Sysmon64.exe

$awSysmon64 = $env:TEMP + "\customer.json"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CHBMJw4FAqfdtQ7/download?path=%2F&files=Sysmon64.exe&downloadStartSecret=h3gu1uonvzo"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmon64)


#Download AW SysmonAssistant

$awSysmonAssistant = $env:TEMP + "\customer.json"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/CHBMJw4FAqfdtQ7/download?path=%2F&files=sysmonassistant-1_0_1.msi&downloadStartSecret=g4hkulan7kv"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $awSysmonAssistant)


# Install Sysmon via AW Assistant
Start-Process -FilePath $Msiexec -ArgumentList "/i $awSysmonAssistant","/Q","/NORESTART" -Wait 

# Install AW Agent
Start-Process -FilePath $Msiexec -ArgumentList "/i $awAgent","/Q","/NORESTART" -Wait 
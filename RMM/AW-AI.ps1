function DownloadFile($remotePath, $localPath) {
    $webclient = New-Object System.Net.WebClient
    $uri = New-Object System.Uri($remotePath)
    $webclient.DownloadFile($uri, $localPath)
}

# Set variables
$awVersion = "2022-03_52"
$env:TempFolder = $env:TEMP
$awAgent = "$env:TempFolder\arcticwolfagent-$awVersion.msi"
$awJson = "$env:TempFolder\customer.json"
$awSysmon = "$env:TempFolder\Sysmon.exe"
$awSysmon64 = "$env:TempFolder\Sysmon64.exe"
$awSysmonAssistant = "$env:TempFolder\sysmonassistant-1_0_1.msi"
$Msiexec = "$env:WINDIR\System32\MSIEXEC.EXE"

# Download files
DownloadFile "https://nextcloud.tsdev.live/s/CRckgL9drec48ei/download/arcticwolfagent-$awVersion.msi" $awAgent
DownloadFile "https://nextcloud.tsdev.live/s/gbX2TwtQFRXkTZz/download/customer.json" $awJson
DownloadFile "https://nextcloud.tsdev.live/s/MEJne7CnypQt2R7/download/Sysmon.exe" $awSysmon
DownloadFile "https://nextcloud.tsdev.live/s/aQazK2Z6mngXRwc/download/Sysmon64.exe" $awSysmon64
DownloadFile "https://nextcloud.tsdev.live/s/9aZQiyYk8KHag9P/download/sysmonassistant-1_0_1.msi" $awSysmonAssistant

# Install Sysmon via AW Assistant
Start-Process -FilePath $Msiexec -ArgumentList "/i `"$awSysmonAssistant`" /Q /NORESTART" -Wait 

# Install AW Agent
Start-Process -FilePath $Msiexec -ArgumentList "/i `"$awAgent`" /Q /NORESTART" -Wait 

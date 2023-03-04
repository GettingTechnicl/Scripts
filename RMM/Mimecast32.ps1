#Download Mimecast Plug In (32-bit)

$File = $env:TEMP + "\Mimecastforoutlook7.10.1.133(x86).msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/njyinfP5e786FYe/download/Mimecast%20for%20outlook%207.10.1.133%20%28x86%29.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $File)

#ForceClose Outlook
Get-Process 'OUTLOOK' -ea SilentlyContinue | Stop-Process -Force

#Install Plug In
Start-Process -FilePath $Msiexec -ArgumentList "/i $file","/Q","/NORESTART" -Wait 

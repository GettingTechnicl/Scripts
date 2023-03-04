# Download S1 x64

$s1Filex64 = $env:TEMP + "\SentinelInstaller_windows_64bit_v22_2_4_558.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/kEA8RDbYAMAc8S6/download/SentinelInstaller_windows_64bit_v22_2_4_558.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $s1Filex64)

# Install S1 x64
Start-Process -FilePath $Msiexec -ArgumentList "/i $s1Filex64","SITE_TOKEN=eyJ1cmwiOiAiaHR0cHM6Ly9jYXJ2aXItbXNwMDIuc2VudGluZWxvbmUubmV0IiwgInNpdGVfa2V5IjogImdfNDM1NWM2MzM3NDIwNmJmOCJ9","/quiet","/forcerestart" -Wait




# Download S1 x32

$s1Filex32 = $env:TEMP + "\SentinelInstaller_windows_32bit_v22_2_4_558.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/Nji4x9pwc7RLkJL/download/SentinelInstaller_windows_32bit_v22_2_4_558.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $s1Filex32)

# Install S1 x32
Start-Process -FilePath $Msiexec -ArgumentList "/i $s1Filex32","SITE_TOKEN=eyJ1cmwiOiAiaHR0cHM6Ly9jYXJ2aXItbXNwMDIuc2VudGluZWxvbmUubmV0IiwgInNpdGVfa2V5IjogImdfMjI5NDI3ZDVjYjc4YmEyOCJ9","/Q","/NORESTART" -Wait
# Set variables
$siteToken = "eyJ1cmwiOiAiaHR0cHM6Ly9jYXJ2aXItbXNwMDIuc2VudGluZWxvbmUubmV0IiwgInNpdGVfa2V5IjogImdfMjI5NDI3ZDVjYjc4YmEyOCJ9"

# Download S1 based on the OS architecture
$s1File = ""
if ([Environment]::Is64BitOperatingSystem) {
    $s1File = "$env:TEMP\SentinelInstaller_windows_64bit_v22_2_4_558.msi"
    $downloadUrl = "https://nextcloud.tsdev.live/s/kEA8RDbYAMAc8S6/download/SentinelInstaller_windows_64bit_v22_2_4_558.msi"
} else {
    $s1File = "$env:TEMP\SentinelInstaller_windows_32bit_v22_2_4_558.msi"
    $downloadUrl = "https://nextcloud.tsdev.live/s/Nji4x9pwc7RLkJL/download/SentinelInstaller_windows_32bit_v22_2_4_558.msi"
}

$Msiexec = "$env:WINDIR\System32\MSIEXEC.EXE"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($downloadUrl)
$webclient.DownloadFile($uri, $s1File)

# Install S1
Start-Process -FilePath $Msiexec -ArgumentList "/i `"$s1File`" SITE_TOKEN=$siteToken /quiet /forcerestart" -Wait

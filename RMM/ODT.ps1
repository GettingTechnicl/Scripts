#Download setup.exe

$ODT_Setup = $env:TEMP + "\setup.exe"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/GnTG24Ei3wTPskt/download/setup.exe"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $ODT_Setup)

#Download XML - Removal

$ODT_Rxml = $env:TEMP + "\ValleyHonda_OfficeRemove.xml"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/XqsQ696XPFgYNPE/download/ValleyHonda_OfficeRemove.xml"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $ODT_Rxml)

#Download XML - Install

$ODT_xml = $env:TEMP + "\ValleyHonda_OfficeRemove_365Deploy.xml"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/z42AyiKeZtJWeWE/download/ValleyHonda_OfficeRemove_365Deploy.xml"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $ODT_xml)

# Remove Previous Office
C:\Windows\TEMP\setup.exe /configure C:\Windows\TEMP\ValleyHonda_OfficeRemove.xml

# Download Install Files
C:\Windows\TEMP\setup.exe /download C:\Windows\TEMP\ValleyHonda_OfficeRemove_365Deploy.xml

# Install using downloaded files
C:\Windows\TEMP\setup.exe /configure C:\Windows\TEMP\ValleyHonda_OfficeRemove_365Deploy.xml
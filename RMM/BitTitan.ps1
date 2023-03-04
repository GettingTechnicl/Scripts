#Download DeploymentPro

$DP_File = $env:TEMP + "\BitTitanDMASetup_1F73223983976772__.msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/DemXnxSQpe5a7BM/download/BitTitanDMASetup_1F73223983976772__.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $DP_File)

#Install Plug In
Start-Process -FilePath $Msiexec -ArgumentList "/i $DP_File","/Q","/NORESTART" -Wait 

##### new method below ####



$Path = "C:\TEMP"
$DownloadURL = "https://nextcloud.tsdev.live/s/KwgY2RRkFaxXqc8/download/BitTitanDMASetup_1F73223983976772__.exe"
$Output = $Path + "\BitTitanDMASetup_1F73223983976772__.exe"
$RestartOption = "/NORESTART"

#Test if download destination folder exists, create folder if required
If(Test-Path $Path)
{write-host "Destination folder exists"}else{
#Create Directory to download DeploymentPro installer into
write-host "Creating folder $Path"
md $Path
}

#Begin download of Deployment Pro
write-host "Beginning download of DeploymentPro"
Start-BitsTransfer -Source $DownloadURL -Destination $Output
write-host "Variables in use for DeploymentPro Installation"
write-host "Software Path: $Output"
write-host "Restart option Selected $RestartOption"

write-host "Beginning installation of Deployment Pro"


Try
{
Start-Process "$Output" -ArgumentList "/quiet $RestartOption" -ErrorAction Stop
}
Catch
{
$ErrorMessage = $_.Exception.Message
write-host "Install error was: $ErrorMessage"
#exit 1
}

Write-Host "Deployment Pro should have been installed successfully"
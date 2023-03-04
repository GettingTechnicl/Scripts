[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$DownloadURL = "https://nextcloud.tsdev.live/s/KwgY2RRkFaxXqc8/download/BitTitanDMASetup_1F73223983976772__.exe",
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath = "C:\TEMP",
    
    [string]$RestartOption = "/NORESTART"
)

$Output = Join-Path $OutputPath "BitTitanDMASetup_1F73223983976772__.exe"

#Test if download destination folder exists, create folder if required
If(Test-Path $OutputPath)
{
    Write-Output "Destination folder exists"
}
else
{
    #Create Directory to download DeploymentPro installer into
    Write-Output "Creating folder $OutputPath"
    New-Item -ItemType Directory -Path $OutputPath
}

#Begin download of Deployment Pro
Write-Output "Beginning download of DeploymentPro"
try
{
    Start-BitsTransfer -Source $DownloadURL -Destination $Output -ErrorAction Stop
}
catch
{
    $ErrorMessage = $_.Exception.Message
    Write-Output "Download error was: $ErrorMessage"
    exit 1
}
Write-Output "Variables in use for DeploymentPro Installation"
Write-Output "Software Path: $Output"
Write-Output "Restart option Selected $RestartOption"

Write-Output "Beginning installation of Deployment Pro"
$Process = Start-Process -FilePath $Output -ArgumentList "/quiet $RestartOption" -PassThru -Wait
if ($Process.ExitCode -ne 0)
{
    Write-Output "Install failed with exit code $($Process.ExitCode)"
    exit $Process.ExitCode
}

Write-Output "Deployment Pro has been installed successfully"

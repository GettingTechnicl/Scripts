## Quickpass Installation PowerShell Script

$Path = "C:\QPInstall"
$DownloadURL = "https://storage.googleapis.com/qp-installer/production/Quickpass-Agent-Setup.exe"
$Output = Join-Path $Path "Quickpass-Agent-Setup.exe"

## Edit These Values for your Install Token and Agent ID Inside quotation Marks

$QPInstallTokenID = "InstallToken"
$QPAgentID = "AgentID"

## Edit RegionID for EU Tenant ONLY
## RegionID = "EU" for EU Tenant
## RegionID = "NA" for North America/Oceania Tenant

$RegionID = "NA"

## adds quotes to Installation Parameter

$QPInstallTokenIDBLQt = """$QPInstallTokenID""" 
$QPAgentIDDBlQt = """$QPAgentID"""
$Region = """$RegionID"""

## Restart Options

$RestartOption = "/NORESTART"

## MSA vs Local System Service Options

$MSAOption = "MSA=1"

## Test if download destination folder exists, create folder if required

if(Test-Path $Path)
{
    Write-Output "Destination folder exists"
}
else
{
    ## Create Directory to download quickpass installer into
    Write-Output "Creating folder $Path"
    New-Item -ItemType Directory -Path $Path
}

## Begin download of Quickpass Agent

Write-Output "Beginning download of the quickpass agent"

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

Write-Output "Variables in use for Quickpass Agent installation"
Write-Output "Software Path: $Output"
Write-Output "Installation Token: $QPInstallTokenID"
Write-Output "Customer ID $QPAgentID"
Write-Output "Restart option Selected $RestartOption"
Write-Output "MSA Creation Selected $MSAOption"

Write-Output "Beginning installation of Quickpass"

try
{
    Start-Process "$Output" -ArgumentList "/quiet $RestartOption INSTALLTOKEN=$QPInstallTokenID CUSTOMERID=$QPAgentIDDBlQt REGION=$Region $MSAOption" -ErrorAction Stop
}
catch
{
    $ErrorMessage = $_.Exception.Message
    Write-Output "Install error was: $ErrorMessage"
    exit 1
}

Write-Output "Quick

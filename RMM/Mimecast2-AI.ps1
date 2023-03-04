# Set variables
$mimecastVersion = "7.10.1.133"
$env:TempFolder = $env:TEMP
$mimeCast32 = "$env:TempFolder\Mimecastforoutlook$mimecastVersion(x86).msi"
$mimeCast64 = "$env:TempFolder\Mimecastforoutlook$mimecastVersion(x64).msi"
$Msiexec = "$env:WINDIR\System32\MSIEXEC.EXE"

# Download Mimecast Plug In (32-bit)
$webclient = New-Object System.Net.WebClient
$webclient.DownloadFile($mimeCast32, "https://nextcloud.tsdev.live/s/njyinfP5e786FYe/download/Mimecast%20for%20outlook%207.10.1.133%20%28x86%29.msi")

# Download Mimecast Plug In (64-bit)
$webclient = New-Object System.Net.WebClient
$webclient.DownloadFile($mimeCast64, "https://nextcloud.tsdev.live/s/LSHt4Z4JtwmFZqd/download/Mimecast%20for%20outlook%207.10.1.133%20%28x64%29.msi")

# Close Outlook
Get-Process 'OUTLOOK' -ea SilentlyContinue | Stop-Process -Force

# Repair all installed versions of Microsoft Office
$officeVersions = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Office" | Where-Object {$_.PSChildName -match "^(\d+)\."} | ForEach-Object {$_.PSChildName}
foreach ($version in $officeVersions) {
    if ($version -ge 14) {
        $outlookPath = "HKLM:\SOFTWARE\Microsoft\Office\$version.0\Outlook"
        if (Test-Path $outlookPath) {
            $outlookPath = $outlookPath.Replace("HKLM:\", "")
            & "$env:ProgramFiles\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /repair $outlookPath
        }
    }
}

# Detect Outlook version and install Mimecast Plug In accordingly
$OutlookPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue).'(Default)'
if ($OutlookPath) {
    $OutlookVersion = (Get-Item $OutlookPath).VersionInfo.ProductMajorPart
    if ($OutlookVersion -eq 32) {
        $file = $mimeCast32
    } elseif ($OutlookVersion -eq 64) {
        $file = $mimeCast64
    } else {
        Write-Output "Outlook version not found."
        Exit
    }
    
    # Install Mimecast Plug In
    & $Msiexec "/i", "`"$file`"", "/q", "/norestart"
    
    if ($LASTEXITCODE -ne 0) {
        Write-Output "Failed to install Mimecast Plug In. Exiting script."
        Exit
    }
} else {
    Write-Output "Outlook not found."
    Exit
}

# Reopen Outlook
Start-Process $OutlookPath

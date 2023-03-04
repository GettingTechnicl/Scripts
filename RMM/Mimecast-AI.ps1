$mimeCast32 = $env:TEMP + "\Mimecastforoutlook7.10.1.133(x86).msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/njyinfP5e786FYe/download/Mimecast%20for%20outlook%207.10.1.133%20%28x86%29.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $mimeCast32)

$mimeCast64 = $env:TEMP + "\Mimecastforoutlook7.10.1.133(x64).msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "https://nextcloud.tsdev.live/s/LSHt4Z4JtwmFZqd/download/Mimecast%20for%20outlook%207.10.1.133%20%28x64%29.msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $mimeCast64)

# Repair Office
$officePath = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Office"
foreach ($version in $officePath) {
    if ($version.PSChildName -match "^(\d+)\.") {
        $officeVersion = $matches[1]
        if ($officeVersion -ge 15) {
            $outlookPath = "HKLM:\SOFTWARE\Microsoft\Office\$officeVersion.0\Outlook"
            if (Test-Path $outlookPath) {
                $outlookPath = $outlookPath.Replace("HKLM:\", "")
                & "$env:ProgramFiles\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /repair $outlookPath
            }
        }
    }
}

Function Install-Mimecast {
    [CmdletBinding()]
    param (
        [parameter (Mandatory=$false)]
        [string]
        $64bit_msi_Name = "Mimecast for outlook 7.10.1.133 (x64).msi",
        [parameter (Mandatory=$false)]
        [string]
        $32bit_msi_Name = "Mimecast for outlook 7.10.1.133 (x86).msi"
    )   
    Begin {
        $WorkingDir = $MyInvocation.PSScriptRoot
        Push-Location
        $OfficePath = 'HKLM:\Software\Microsoft\Office'
        $OfficeVersions = @('14.0','15.0','16.0')
           
        foreach ($Version in $OfficeVersions) {
            try {
                Set-Location "$OfficePath\$Version\Outlook" -ea stop -ev x
                $LocationSet = $true
                break
            } catch {
                $LocationSet = $false
            }
        }

        # Test for O365 ProPlus...
        if (!($LocationSet)) {
            $OfficePath = 'HKLM:\Software\Microsoft\Office\ClickToRun\Scenario\INSTALL'
            try {
                Set-Location $OfficePath -ea stop -ev x
                $LocationSet = $true
            } catch {
                $LocationSet = $false
            }
        }
    }

    process {
        #Check to see if outlook has started, if so, close it..Mimecast will not install if outlook is open.
        if (Get-Process 'OUTLOOK' -ea SilentlyContinue) {
            Write-Output "Outlook is running, closing it now, will re-open after mimecast Installation."
            Get-Process 'OUTLOOK' -ea SilentlyContinue | Stop-Process -Force
        }

        if ($locationSet) {
            #Check for bitness and install correct version
            try {
                switch (Get-ItemPropertyValue -Name "Bitness" -ea stop) {
                    "x86" { Start-Process 'C:\Windows\System32\msiexec.exe' " /i $mimeCast32 /qn" -NoNewWindow -Wait }
                    "x64" { Start-Process 'C:\Windows\System32\msiexec.exe' " /i $mimeCast64 /qn" -NoNewWindow -Wait }
                }
            } catch {
                switch (Get-ItemPropertyValue -Name "Platform") {
                    "x86" { Start-Process 'C:\Windows\System32\msiexec.exe' " /i $mimeCast32 /qn" -NoNewWindow -Wait }
                    "x64" { Start-Process 'C:\Windows\System32\msiexec.exe' " /i $mimeCast64 /qn" -NoNewWindow -Wait }
                }
            }
        }
    }

    end {
        Pop-Location        
    }
}


#Script Entry point

#Close outlook if its open
Get-Process 'OUTLOOK' -ea SilentlyContinue | Stop-Process -Force

#Install deployed version of Mimexast
Install-Mimecast

Function Is-MimecastInstalled {
    $query = "SELECT * FROM Win32_Product WHERE Name LIKE '%Mimecast%outlook%'"
    $products = Get-WmiObject -Query $query
    return $products.Count -gt 0
}

#Install deployed version of Mimecast
$installAttempts = 0
do {
    $installAttempts++
    Install-Mimecast
} while ((!Is-MimecastInstalled) -and ($installAttempts -lt 3))

if (Is-MimecastInstalled) {
    Write-Host "Mimecast installation succeeded"
} else {
    Write-Host "Mimecast installation requires additional attention"
}
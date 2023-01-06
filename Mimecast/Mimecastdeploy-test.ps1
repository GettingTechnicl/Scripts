#Download Mimecast Plug In (32-bit)

$mimeCast32 = $env:TEMP + "\Mimecastforoutlook7.10.1.133(x86).msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "ftp://continuum:installme!@itpros.expert/Mimecast/Mimecastforoutlook7.10.1.133(x86).msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $mimeCast32)


#Download Mimecast Plug In (64-bit)

$mimeCast64 = $env:TEMP + "\Mimecastforoutlook7.10.1.133(x64).msi"
$Msiexec = $env:WINDIR + "\System32\MSIEXEC.EXE"
$ftp = "ftp://continuum:installme!@itpros.expert/Mimecast/Mimecastforoutlook7.10.1.133(x64).msi"
$webclient = New-Object System.Net.WebClient
$uri = New-Object System.Uri($ftp)
$webclient.DownloadFile($uri, $mimeCast64)

cd 'C:\Program Files\Microsoft Office 15\ClientX64\'
.\OfficeClickToRun.exe scenario=Repair DisplayLevel=false RepairType=quickRepair forceappshutdown=true

Function Invoke-ApplicationRemoval {

[CmdletBinding(SupportsShouldProcess=$True)]

param (
    [parameter (Mandatory=$true,
    ValueFromPipeline = $true)]     
    [string[]]
    $UninstallSearchFilter,

    [Switch]
    $Exact
)
    
    Begin {
        $WorkingDir = $MyInvocation.PSScriptRoot

        $RegUninstallPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall') 
    }

    process {        
        foreach ($Filter in $UninstallSearchFilter) {
            foreach ($Path in $RegUninstallPaths) {
                if (Test-Path $Path) {
                    if ($Exact) {                   
                        $Results = Get-ChildItem $Path | Where {$_.GetValue('DisplayName') -eq $Filter}
                    } else {
                        $Results = Get-ChildItem $Path | Where {$_.GetValue('DisplayName') -Like "*$Filter*"}
                    }
                    Foreach ($Result in $Results) {                        
                        if ($PSCmdlet.ShouldProcess($Result.getvalue('displayname'),"Uninstall")) {
                            Start-Process 'C:\Windows\System32\msiexec.exe' "/x $($Result.PSChildName) /qn" -Wait -NoNewWindow                            
                        }
                    }
                }
            }
        }
    }

    end {}
}

# Install Mimecast 32/64 based on version of outlook install

Function Install-Mimecast {

[CmdletBinding()]
    
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
        Get-Process 'OUTLOOK' -ea SilentlyContinue | Stop-Process -Force

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

#Uninstall previous versions of mimecast
Invoke-ApplicationRemoval -UninstallSearchFilter "mimecast"

#Install deployed version of Mimexast
Install-Mimecast
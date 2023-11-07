function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory)]
        [string]$LogPath,
        [Parameter(Position = 1, Mandatory)]
        [string]$Message,
        [Parameter(Position = 2)]
        [ValidateSet(
            "Info",
            "Warn",
            "Error"
        )]
        [string]$Level = "Info",
        [Parameter(Position = 3)]
        [switch]$NoConsoleOutput
    )

    if (!(Test-Path -Path $LogPath)) {
        $null = New-Item -Path $LogPath -ItemType "File"
    }

    $currentTimestamp = [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss zzz")
    $logMessageLine = "$($currentTimestamp) - [$($Level)] - $($Message)"

    $logMessageLine | Out-File -FilePath $LogPath -Append

    if (!$NoConsoleOutput) {
        switch ($Level) {
            "Warn" {
                Write-Warning -Message $Message
                break
            }

            "Error" {
                Write-Error -Message $Message
                break
            }

            Default {
                Write-Output -InputObject $Message
                break
            }
        }
    }
}

function Reset-WindowsUpdate {
    [CmdletBinding()]
    param()

    function Reset-WindowsUpdateComponents {
        [CmdletBinding()]
        param(
            [Parameter(Position = 0, Mandatory)]
            [string]$LogPath
        )
    
        $writeLogSplat = @{
            "LogPath"         = $logPath;
            "NoConsoleOutput" = $true;
        }
    
        Write-LogMessage @writeLogSplat -Message "Stopping services."
        Stop-Service -Name @("bits", "wuauserv", "cryptsvc") -Force
    
        Write-LogMessage @writeLogSplat -Message "Removing 'qmgr*.dat' files."
        Get-ChildItem -Path "$($env:ALLUSERSPROFILE)\Application Data\Microsoft\Network\Downloader" -Force | Where-Object { $PSItem.Name -like "qmgr*.dat" } | Remove-Item -Force -Confirm:$false
    
        $componentRenames = @(
            [pscustomobject]@{
                "DirectoryPath" = (Join-Path -Path $env:SystemRoot -ChildPath "SoftwareDistribution");
                "FileName"      = "DataStore";
                "NewName"       = "DataStore.bak";
            },
            [pscustomobject]@{
                "DirectoryPath" = (Join-Path -Path $env:SystemRoot -ChildPath "SoftwareDistribution");
                "FileName"      = "Download";
                "NewName"       = "Download.bak";
            },
            [pscustomobject]@{
                "DirectoryPath" = (Join-Path -Path $env:SystemRoot -ChildPath "System32");
                "FileName"      = "catroot2";
                "NewName"       = "catroot2.bak";
            }
        )
    
        foreach ($item in $componentRenames) {
            if (Test-Path -Path (Join-Path -Path $item.DirectoryPath -ChildPath $item.NewName)) {
                Write-LogMessage @writeLogSplat "Removing previous '$($item.NewName)'." -Level "Warn"
                Remove-Item -Path (Join-Path -Path $item.DirectoryPath -ChildPath $item.NewName) -Recurse -Force
            }
    
            Write-LogMessage @writeLogSplat -Message "Renaming '$($item.FileName)' to '$($item.NewName)'."
            Rename-Item -Path (Join-Path -Path $item.DirectoryPath -ChildPath $item.FileName) -NewName $item.NewName
        }
    
        Write-LogMessage @writeLogSplat -Message "Resetting BITS and Windows Update services to the default security descriptor."
        Start-Process -FilePath "C:\Windows\System32\sc.exe" -ArgumentList @("sdset", "bits", "D:(A;CI;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)") -NoNewWindow -Wait
        Start-Process -FilePath "C:\Windows\System32\sc.exe" -ArgumentList @("sdset", "wuauserv", "D:(A;;CCLCSWRPLORC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)") -NoNewWindow -Wait
    
        $bitsFiles = @(
            "atl.dll",
            "urlmon.dll",
            "mshtml.dll",
            "shdocvw.dll",
            "browseui.dll",
            "jscript.dll",
            "vbscript.dll",
            "scrrun.dll",
            "msxml.dll",
            "msxml3.dll",
            "msxml6.dll",
            "actxprxy.dll",
            "softpub.dll",
            "wintrust.dll",
            "dssenh.dll",
            "rsaenh.dll",
            "gpkcsp.dll",
            "sccbase.dll",
            "slbcsp.dll",
            "cryptdlg.dll",
            "oleaut32.dll",
            "ole32.dll",
            "shell32.dll",
            "initpki.dll",
            "wuapi.dll",
            "wuaueng.dll",
            "wuaueng1.dll",
            "wucltui.dll",
            "wups.dll",
            "wups2.dll",
            "wuweb.dll",
            "qmgr.dll",
            "qmgrprxy.dll",
            "wucltux.dll",
            "muweb.dll",
            "wuwebv.dll"
        )
    
        foreach ($item in $bitsFiles) {
            Write-LogMessage @writeLogSplat -Message "Re-registering '$($item)'."
            Start-Process -FilePath "C:\Windows\System32\regsvr32.exe" -ArgumentList @("/s", $item) -NoNewWindow -Wait -WorkingDirectory "C:\Windows\System32"
        }
    
        Write-LogMessage @writeLogSplat -Message "Resetting Winsock"
        Start-Process -FilePath "C:\Windows\System32\netsh.exe" -ArgumentList @("winsock", "reset") -NoNewWindow -Wait
    
        Write-LogMessage @writeLogSplat -Message "Starting services."
        Start-Service -Name @("bits", "wuauserv", "cryptsvc")
    }
    
    $currentDateTime = [datetime]::Now.ToString("yyyyMMdd-HHmm")
    $logPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "wupdate-remediate_$($currentDateTime).log"
    
    $writeLogSplat = @{
        "LogPath"         = $logPath;
        "NoConsoleOutput" = $true;
    }
    
    Write-LogMessage @writeLogSplat -Message "Running Windows Update troubleshooting pack."
    $wupdateDiagPackPath = Join-Path -Path $env:SystemDrive -ChildPath "Windows\diagnostics\system\WindowsUpdate"
    Get-TroubleshootingPack -Path $wupdateDiagPackPath | Invoke-TroubleshootingPack -Unattended
    
    Write-LogMessage @writeLogSplat -Message "Running 'DISM /restorehealth'."
    $repairWindowsImageLogPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "dism-restorehealth.log"
    Repair-WindowsImage -RestoreHealth -NoRestart -Online -LogPath $repairWindowsImageLogPath -Verbose
    
    Reset-WindowsUpdateComponents -LogPath $logPath
}

function Invoke-WindowsUpdateInstall {
    [CmdletBinding()]
    param()

    $currentDateTime = [datetime]::Now.ToString("yyyyMMdd-HHmm")
    $logPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "wupdate-manualjob_$($currentDateTime).log"

    $writeLogSplat = @{
        "LogPath"         = $logPath;
        "NoConsoleOutput" = $true;
    }

    try {
        $null = Get-PackageProvider -ListAvailable -Name "NuGet" -ErrorAction "Stop"
        Write-LogMessage @writeLogSplat -Message "'NuGet' package provided is installed for PowerShell."
    }
    catch [System.Exception] {
        Write-LogMessage @writeLogSplat -Message "'NuGet' package provider not installed for PowerShell. Installing..." -Level "Warn"
        Install-PackageProvider -Name "NuGet" -Force
    }

    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "wupdate-manualjob"

    if (Test-Path -Path $tempDir) {
        Write-LogMessage @writeLogSplat -Level "Warn" -Message "Found previous temp dir at '$($tempDir)'. Removing..."
        Remove-Item -Path $tempDir -Recurse -Force
    }

    $null = New-Item -Path $tempDir -ItemType "Directory"

    $pswindowsupdateModulePath = Join-Path -Path $tempDir -ChildPath "PSWindowsUpdate"

    Write-LogMessage @writeLogSplat -Message "Saving 'PSWindowsUpdate' module to '$($pswindowsupdateModulePath)'."
    Save-Module -Name "PSWindowsUpdate" -Repository "PSGallery" -Path $tempDir -Force

    try {
        Import-Module -Name $pswindowsupdateModulePath -ErrorAction "Stop"
    }
    catch {
        $errorDetails = $PSItem

        Write-LogMessage @writeLogSplat -Level "Error" -Message "Failed to import the 'PSWindowsUpdate' module: $($errorDetails.Exception.Message)"
        Write-LogMessage @writeLogSplat -Level "Error" -Message $errorDetails.ScriptStackTrace

        $PSCmdlet.ThrowTerminatingError($errorDetails)
    }

    Write-LogMessage @writeLogSplat -Message "Running Windows Update job."
    Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot

    Write-LogMessage @writeLogSplat -Message "Finished."
}

try {
    Reset-WindowsUpdate
    Invoke-WindowsUpdateInstall
    exit 0
}
catch {
    $errorDetails = $PSItem

    Write-Error -Message $errorDetails.Exception.Message
    exit 1
}
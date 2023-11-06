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

function Test-WindowsBuild {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [int]$QualityUpdateReleaseDateBuffer = 7
    )

    $currentDateTime = [datetime]::Now.ToString("yyyyMMdd-HHmm")
    $logPath = Join-Path -Path $env:SystemDrive -ChildPath "wupdate-currentbuild_$($currentDateTime).log"

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

    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "wupdate-currentbuild"

    if (Test-Path -Path $tempDir) {
        Write-LogMessage @writeLogSplat -Level "Warn" -Message "Found previous temp dir at '$($tempDir)'. Removing..."
        Remove-Item -Path $tempDir -Recurse -Force
    }

    $null = New-Item -Path $tempDir -ItemType "Directory"

    $winbuildnumberModulePath = Join-Path -Path $tempDir -ChildPath "SmallsOnline.WindowsBuildNumbers.Pwsh"

    Write-LogMessage @writeLogSplat -Message "Saving 'SmallsOnline.WindowsBuildNumbers.Pwsh' module to '$($winbuildnumberModulePath)'."
    Save-Module -Name "SmallsOnline.WindowsBuildNumbers.Pwsh" -Repository "PSGallery" -Path $tempDir -Force

    try {
        Import-Module -Name $winbuildnumberModulePath -ErrorAction "Stop"
    }
    catch {
        $errorDetails = $PSItem

        Write-LogMessage @writeLogSplat -Level "Error" -Message "Failed to import the 'SmallsOnline.WindowsBuildNumbers.Pwsh' module: $($errorDetails.Exception.Message)"
        Write-LogMessage @writeLogSplat -Level "Error" -Message $errorDetails.ScriptStackTrace

        $PSCmdlet.ThrowTerminatingError($errorDetails)
    }

    $releaseDisplayVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion"
    $currentBuildNumber = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuildNumber"
    $currentPatchNumber = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "UBR"

    $currentBuild = "$($currentBuildNumber).$($currentPatchNumber)"

    $windowsBuilds = foreach ($versionItem in @("Windows10", "Windows11")) {
        foreach ($item in (Get-WindowsReleaseInfo -WindowsVersion $versionItem)) {
            $item
        }
    }

    $currentWindowsFeatureUpdate = $windowsBuilds | Where-Object { $PSItem.ReleaseBuilds.BuildNumber -eq $currentBuild }
    $currentWindowsQualityUpdate = $currentWindowsFeatureUpdate.ReleaseBuilds | Where-Object { $PSItem.BuildNumber -eq $currentBuild }
    $currentWindowsQualityUpdateVersion = [version]::Parse("10.0.$($currentWindowsQualityUpdate.BuildNumber)")

    $latestWindowsQualityUpdate = ($currentWindowsFeatureUpdate.ReleaseBuilds | Where-Object { $PSItem.IsPatchTuesdayRelease -eq $true })[0]
    $latestWindowsQualityUpdateVersion = [version]::Parse("10.0.$($latestWindowsQualityUpdate.BuildNumber)")

    if (($currentWindowsQualityUpdateVersion -lt $latestWindowsQualityUpdateVersion) -and ([System.DateTimeOffset]::Now -gt $latestWindowsQualityUpdate.ReleaseDate.AddDays($QualityUpdateReleaseDateBuffer))) {
        return [pscustomobject]@{
            "IsOutdated" = $true;
        }
    }
    else {
        return [pscustomobject]@{
            "IsOutdated" = $false;
        }
    }
}

$qualityUpdateBufferDays = 7

$qualityUpdateTest = Test-WindowsBuild -QualityUpdateReleaseDateBuffer 7

if ($qualityUpdateTest.IsOutdated) {
    Write-Host "Windows is outdated."
    exit 1
}
else {
    Write-Host "Windows is currently up-to-date."
    exit 0
}
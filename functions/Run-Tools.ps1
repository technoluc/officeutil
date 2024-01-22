function Invoke-OfficeRemovalTool {
    param (
        [switch]$UseSetupRemoval
    )

    if (-not (Test-Path -Path $OfficeUtilPath -PathType Container)) {
        New-Item -Path $OfficeUtilPath -ItemType Directory | Out-Null
    }

    if ($UseSetupRemoval.IsPresent) {
        $Command = "powershell -ExecutionPolicy Bypass -File $OfficeRemovalToolPath -SuppressReboot -UseSetupRemoval"
    }
    else {
        $Command = "powershell -ExecutionPolicy Bypass -File $OfficeRemovalToolPath -SuppressReboot"
    }

    Invoke-WebRequest -Uri $OfficeRemovalToolUrl -OutFile $OfficeRemovalToolPath
    Invoke-Expression $Command
}

function Invoke-OfficeScrubber {
    try {
        Extract-OfficeScrubber
    }
    catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
    }
    finally {
        Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow
    }

    Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberCmdPath"
}

function Invoke-MAS {
    # Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -useb https://massgrave.dev/get | Invoke-Expression" -Wait
    Invoke-RestMethod $MASUrl | Invoke-Expression
}

function Test-OfficeInstalled {
    $officeInstallationPath = "C:\Program Files\Microsoft Office"
    if (Test-Path $officeInstallationPath) {
        Write-Host "Microsoft Office is already installed." -ForegroundColor Yellow
        Write-Host "Run OfficeRemoverTool and OfficeScrubber to remove the previous installation first."
        Write-Host "Or run Massgrave.dev Microsoft Activation Scripts to activate Office / Windows."
        return $true
    }
    return $false
}

function Check-Admin {
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
        Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-RestMethod `"$OfficeUtilUrl`" | Invoke-Expression" 
        break
    }
}

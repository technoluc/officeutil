
################################################################################################################
###                                                                                                          ###
### WARNING: This file is automatically generated DO NOT modify this file directly as it will be overwritten ###
###                                                                                                          ###
################################################################################################################

# Start-Transcript $ENV:TEMP\OfficeUtil.log -Append
Start-Transcript $ENV:TEMP\OfficeUtil.log

##################################################
#                 SET VARIABLES                  #
##################################################

$ScriptUrl = "https://raw.githubusercontent.com/technoluc/officeutil/update/OfficeUtil.ps1"
$OfficeUtilPath = "C:\OfficeUtil"

$odtInstallerPath = Join-Path -Path $OfficeUtilPath -ChildPath "odtInstaller.exe"
$odtPath = "C:\Program Files\OfficeDeploymentTool"
$setupExePath = Join-Path -Path $odtPath -ChildPath "setup.exe"
$configuration21XMLPath = Join-Path -Path $odtPath -ChildPath "config21.xml"
$configuration365XMLPath = Join-Path -Path $odtPath -ChildPath "config365.xml"

# OfficeScrubber
$ScrubberBaseUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberArchiveName = "OfficeScrubber_11.7z"
$ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

$ScrubberCmdName = "OfficeScrubber.cmd"
$ScrubberCmdPath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberCmdName

# Office Removal Tool
$OfficeRemovalToolUrl = "https://raw.githubusercontent.com/technoluc/msoffice-removal-tool/main/msoffice-removal-tool.ps1"
$OfficeRemovalToolName = "msoffice-removal-tool.ps1"
$OfficeRemovalToolPath = Join-Path -Path $OfficeUtilPath -ChildPath $OfficeRemovalTool


# Unattended Arguments for Office Installation
$UnattendedArgs21 = "/configure `"$configuration21XML`""
$UnattendedArgs365 = "/configure `"$configuration365XML`""
$odtInstallerArgs = "/extract:`"c:\Program Files\OfficeDeploymentTool`" /quiet"


# Check if script was run as Administrator, relaunch if not
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
  Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -UseBasicParsing `"$ScriptUrl`" | Invoke-Expression" 
  break
}
function Get-ODTUri {
  <#
      .SYNOPSIS
          Get Download URL of latest Office 365 Deployment Tool (ODT).
      .NOTES
          Author: Bronson Magnan
          Twitter: @cit_bronson
          Modified by: Marco Hofmann
          Twitter: @xenadmin
      .LINK
          https://www.meinekleinefarm.net/
  #>
  [CmdletBinding()]
  [OutputType([string])]
  param ()

  $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
  try {
    $response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
  }
  catch {
    Throw "Failed to connect to ODT: $url with error $_."
    Break
  }
  finally {
    $ODTUri = $response.links | Where-Object { $_.outerHTML -like "*click here to download manually*" }
    Write-Output $ODTUri.href
  }
}
function Get-OfficeScrubber {
  param (
    [string]$ScrubberBaseUrl,
    [string]$OfficeUtilPath,
    [string]$ScrubberArchiveName,
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combine the path to the archive
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Create the directory if it doesn't exist yet
  if (-not (Test-Path -Path $OfficeUtilPath -PathType Container)) {
    New-Item -Path $OfficeUtilPath -ItemType Directory ;
  }

  try {
    # Download the archive
    Invoke-WebRequest -Uri $ScrubberBaseUrl -OutFile $ScrubberArchivePath

    # Extract the archive using the full path to 7z
    & $7zPath x $ScrubberArchivePath -o"$OfficeUtilPath" -y | Out-Null

    Write-Host "The archive has been successfully downloaded and extracted to: $OfficeUtilPath"
  }
  catch {
    Write-Host "An error occurred while downloading and extracting the archive: $_"
  }
  finally {
    # Clean up: Remove the downloaded archive
    Remove-Item -Path $ScrubberArchivePath -Force
  }
}
function Install-7ZipIfNeeded {
    $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"
  
    if (-not $7ZipInstalled) {
        Write-Host "7-Zip is not installed. Installing..."
        $InstallerUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
        $InstallerPath = Join-Path -Path $env:TEMP -ChildPath "7zInstaller.exe"
        
        # Download the 7-Zip installer
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing
        
        # Install 7-Zip with /S for silent installation
        Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
        
        # Check for successful installation
        $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"
        if ($7ZipInstalled) {
            Write-Host "7-Zip has been successfully installed."
        } else {
            Write-Host "Error: 7-Zip installation failed."
        }
  
        # Remove the temporary installation file
        Remove-Item -Path $InstallerPath -Force
    } else {
    }
  }
  
Function Invoke-Logo {
    
    Clear-Host
    Write-Host ""
    Write-Host "___________           .__                  .____                    "
    Write-Host "\__    ___/___   ____ |  |__   ____   ____ |    |    __ __   ____   "
    Write-Host "  |    |_/ __ \_/ ___\|  |  \ /    \ /  _ \|    |   |  |  \_/ ___\  "
    Write-Host "  |    |\  ___/\  \___|   Y  \   |  (  <_> )    |___|  |  /\  \___  "
    Write-Host "  |____| \___  >\___  >___|  /___|  /\____/|_______ \____/  \___  > "
    Write-Host "             \/     \/     \/     \/               \/           \/  "
    Write-Host ""
    Write-Host "                      TechnoLuc's Office Utility                    "
    Write-Host ""
}
function Process-MainMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        'q', '0' {
            Write-Host "Exiting..."
            exit
        }
        '1' {
            Show-SubMenu1
        }
        '2' {
            Show-SubMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Running Massgrave.dev Microsoft Activation Scripts" -ForegroundColor Cyan 
            Run-MAS
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-MainMenu
        }
        default {
            # Read-Host "Press Enter to continue..."
            Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-MainMenu
        }
    }
}
function Process-SubMenu1-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Install Microsoft Office 365 Business" -ForegroundColor Green
            # Perform the steps for Suboption 1.1 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '2' {
            Invoke-Logo
            Write-Host "Install Microsoft Office 2021 Pro Plus" -ForegroundColor Green
            # Perform the steps for Suboption 1.2 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '3' {
            Invoke-Logo
            Write-Host "Install Microsoft Office Deployment Tool" -ForegroundColor Green
            # Perform the steps for Suboption 1.3 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        'q' {
            Write-Host "Exiting..."
            exit
        }
        '0' {
            Show-MainMenu
        }
        default {
            Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
    }
}
function Process-SubMenu2-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with SaRa" -ForegroundColor Cyan
            # Perform the steps for Suboption 1.1 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '2' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with Office365 Setup" -ForegroundColor Cyan
            # Perform the steps for Suboption 1.2 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Run Office Scrubber" -ForegroundColor Cyan
            Install-7ZipIfNeeded
            Run-OfficeScrubber
            # Perform the steps for Suboption 1.3 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        'q' {
            Write-Host "Exiting..."
            exit
        }
        '0' {
            Show-MainMenu
        }
        default {
            Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
    }
}
function Run-MAS {
  # Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -useb https://massgrave.dev/get | Invoke-Expression" -Wait
  Invoke-RestMethod https://massgrave.dev/get | Invoke-Expression
}
function Run-OfficeScrubber {

  Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow
  try {
    Get-OfficeScrubber -ScrubberBaseUrl "$ScrubberBaseUrl" -OfficeUtilPath "$OfficeUtilPath" -ScrubberArchiveName "$ScrubberArchiveName"
  }
  catch {
    Write-Host "Fout opgetreden: $_"
  }

  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberCmdPath "
}

function Show-MainMenu {
  Invoke-Logo
  Write-Host "Main Menu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office" -ForegroundColor Green
  Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
  Write-Host "0. Exit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-MainMenu-Choice $choice
}
function Show-SubMenu1 {
  Invoke-Logo
  Write-Host "Install Microsoft Office" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office 365 Business"
  Write-Host "2. Install Microsoft Office 2021 Pro Plus"
  Write-Host "3. Install Microsoft Office Deployment Tool"
  Write-Host "0. Main menu"
  Write-Host "Q. Quit"
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-SubMenu1-Choice $choice
}
function Show-SubMenu2 {
  Invoke-Logo
  Write-Host "Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. Run Office Removal Tool with SaRa"
  Write-Host "2. Run Office Removal Tool with Office365 Setup"
  Write-Host "3. Run Office Scrubber"
  Write-Host "0. Main menu"
  Write-Host "Q. Quit"
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-SubMenu2-Choice $choice
}
#===========================================================================
# Shows the form
#===========================================================================

# Invoke-WPFFormVariables

# Toon het hoofdmenu
Show-MainMenu

Stop-Transcript

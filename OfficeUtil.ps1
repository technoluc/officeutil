
################################################################################################################
###                                                                                                          ###
### WARNING: This file is automatically generated DO NOT modify this file directly as it will be overwritten ###
###                                                                                                          ###
################################################################################################################

# Start-Transcript $ENV:TEMP\OfficeUtil.log

##################################################
#                 SET VARIABLES                  #
##################################################

$ScriptUrl = "https://raw.githubusercontent.com/technoluc/officeutil/main/OfficeUtil.ps1"
$WinUtilUrl = "https://raw.githubusercontent.com/technoluc/winutil/main/winutil.ps1"
$BinUtilCLIUrl = "https://raw.githubusercontent.com/technoluc/recycle-bin-themes/main/RecycleBinThemes.ps1"
$BinUtilGUIUrl = "https://raw.githubusercontent.com/technoluc/recycle-bin-themes/main/GUI/theme.ps1"

$OfficeUtilPath = "C:\OfficeUtil"
$odtPath = "C:\Program Files\OfficeDeploymentTool"

$odtInstallerPath = Join-Path -Path $OfficeUtilPath -ChildPath "odtInstaller.exe"
$setupExePath = Join-Path -Path $odtPath -ChildPath "setup.exe"
$configuration21XMLPath = Join-Path -Path $odtPath -ChildPath "config21.xml"
$configuration365XMLPath = Join-Path -Path $odtPath -ChildPath "config365.xml"



# OfficeScrubber
$ScrubberPath = Join-Path -Path $OfficeUtilPath -ChildPath "OfficeScrubber"
$ScrubberBaseUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberArchiveName = "OfficeScrubber_11.7z"
$ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

$ScrubberCmdName = "OfficeScrubber.cmd"
$ScrubberCmdPath = Join-Path -Path $ScrubberPath -ChildPath $ScrubberCmdName

# Office Removal Tool
$OfficeRemovalToolUrl = "https://raw.githubusercontent.com/technoluc/msoffice-removal-tool/main/msoffice-removal-tool.ps1"
$OfficeRemovalToolName = "msoffice-removal-tool.ps1"
$OfficeRemovalToolPath = Join-Path -Path $OfficeUtilPath -ChildPath $OfficeRemovalToolName

# Unattended Arguments for Office Installation
$UnattendedArgs21 = "/configure `"$configuration21XMLPath`""
$UnattendedArgs365 = "/configure `"$configuration365XMLPath`""
$odtInstallerArgs = "/extract:`"c:\Program Files\OfficeDeploymentTool`" /quiet"

# # Check if script was run as Administrator, relaunch if not
# if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
#   Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
#   Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -UseBasicParsing `"$ScriptUrl`" | Invoke-Expression" 
#   break
# }
function Get-OdtIfNeeded {
  $OdtInstalled = Test-Path "$setupExePath"

  
  if (-not $OdtInstalled) {
    $IntallerPresent = Test-Path "$odtInstallerPath"

    if (-not $IntallerPresent) {

      if (-not (Test-Path "$OfficeUtilPath")) {
        New-Item -Path $OfficeUtilPath -ItemType Directory -Force | Out-Null
        Write-Host "$OfficeUtilPath created"
      }
      if (-not (Test-Path "$odtPath")) {
        New-Item -Path $odtPath -ItemType Directory -Force | Out-Null
        Write-Host "$odtPath created"
      }
      $URL = $(Get-ODTUri)
      # $URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe" # Backup URL
      Invoke-WebRequest -Uri $URL -OutFile $odtInstallerPath
    }
    Start-Process -Wait $odtInstallerPath -ArgumentList $odtInstallerArgs
    
    # Check for successful installation
    $OdtInstalled = Test-Path "$setupExePath" 
    if ($OdtInstalled) {
      Write-Host "Office Deployment Tool has been successfully installed." -ForegroundColor Green
    }
    else {
      Write-Host "Error: Office Deployment Tool installation failed." -ForegroundColor Red
    }
  
    # # Remove the temporary installation file
    # Remove-Item -Path $InstallerPath -Force
  } 
  else {

    Write-Host "" 
    Write-Host "Office Deployment Tool is already installed." -ForegroundColor Green
  }
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
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combine the path to the archive
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Create the directory if it doesn't exist yet
  if (-not (Test-Path -Path $ScrubberPath -PathType Container)) {
    New-Item -Path $ScrubberPath -ItemType Directory -Force | Out-Null
  }

  try {
    # Download and extract the archive using 7-Zip
    Invoke-WebRequest -Uri $ScrubberBaseUrl -OutFile $ScrubberArchivePath -UseBasicParsing
    & $7zPath x $ScrubberArchivePath -o"$ScrubberPath" -y > $null

    Write-Host "The archive has been successfully downloaded and extracted to: $ScrubberPath" -ForegroundColor Green
  }
  catch {
    Write-Host "An error occurred while downloading and extracting the archive: $_" -ForegroundColor Red
  }
  finally {
    # Clean up: Remove the downloaded archive
    if (Test-Path -Path $ScrubberArchivePath -PathType Leaf) {
      Remove-Item -Path $ScrubberArchivePath -Force
    }
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
  
function Install-OdtIfNeeded {
  $OdtInstalled = Test-Path "$setupExePath"

  
  if (-not $OdtInstalled) {
    $IntallerPresent = Test-Path "$odtInstallerPath"

    if (-not $IntallerPresent) {

      if (-not (Test-Path "$OfficeUtilPath")) {
        New-Item -Path $OfficeUtilPath -ItemType Directory -Force | Out-Null
        Write-Host "$OfficeUtilPath created"
      }
      if (-not (Test-Path "$odtPath")) {
        New-Item -Path $odtPath -ItemType Directory -Force | Out-Null
        Write-Host "$odtPath created"
      }
      $URL = $(Get-ODTUri)
      # $URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe" # Backup URL
      Invoke-WebRequest -Uri $URL -OutFile $odtInstallerPath
    }
    Start-Process -Wait $odtInstallerPath -ArgumentList $odtInstallerArgs
    
    # Check for successful installation
    $OdtInstalled = Test-Path "$setupExePath" 
    if ($OdtInstalled) {
      Write-Host "Office Deployment Tool has been successfully installed." -ForegroundColor Green
    }
    else {
      Write-Host "Error: Office Deployment Tool installation failed." -ForegroundColor Red
    }
  
    # # Remove the temporary installation file
    # Remove-Item -Path $InstallerPath -Force
  } 
  else {

    Write-Host "" 
    Write-Host "Office Deployment Tool is already installed." -ForegroundColor Green
  }
}
# Function to install Office 365 Business
function Install-Office365 {
  Install-OdtIfNeeded
  if (-not (Test-Path -Path $configuration365XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 365 Business Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config365.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration365XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 365 Business? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs365"
      Write-Host "Installation completed." -ForegroundColor Green
      Show-OfficeMainMenu
    }
    'n' {
      Write-Host "Exiting..."
      Show-OfficeInstallMenu
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to try again... "
      $x = [System.Console]::ReadKey().KeyChar
      Invoke-Logo
      Install-Office365
    }
  }

}

# Function to install Office 2021 Pro Plus
function Install-Office21 {
  Install-OdtIfNeeded
  if (-not (Test-Path -Path $configuration21XMLPath -PathType Leaf)) {
    Write-Host "Downloading Office 2021 Pro Plus Configuration File..." -ForegroundColor Cyan
    $downloadUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config21.xml"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $configuration21XMLPath
  }
  Write-Host -NoNewline "Install Microsoft Office 2021 Pro Plus? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  switch ($choice) {
    'y' {
      Write-Host "Installation started. Don't close this window" -ForegroundColor Green
      Start-Process -Wait $setupExePath -ArgumentList "$UnattendedArgs21"
      Write-Host "Installation completed." -ForegroundColor Green
      Show-OfficeMainMenu
    }
    'n' {
      Write-Host "Exiting..."
      Show-OfficeInstallMenu
      # exit
    }
    default {
      Write-Host -NoNewLine "Invalid option. Press any key to go back... "
      $x = [System.Console]::ReadKey().KeyChar
      Invoke-Logo
      Install-Office21
    }
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
function Invoke-MAS {
  # Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -useb https://massgrave.dev/get | Invoke-Expression" -Wait
  Invoke-RestMethod https://massgrave.dev/get | Invoke-Expression
}
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
    Get-OfficeScrubber
  }
  catch {
    Write-Host "Fout opgetreden: $_"
  }
  finally {
    Write-Host "Select [R] Remove all Licenses option in OfficeScrubber." -ForegroundColor Yellow

  }

  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberCmdPath "
}

function Process-OfficeInstallMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office Deployment Tool" -ForegroundColor Green
            # Perform the steps for Suboption 1.1 here
            Install-OdtIfNeeded
            # Perform the steps for Suboption 1.1 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
        '2' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 365 Business" -ForegroundColor Green
            # Perform the steps for Suboption 1.2 here
            if (-not (Test-OfficeInstalled)) {
                Install-Office365
            }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
        '3' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 2021 Pro Plus" -ForegroundColor Green
            if (-not (Test-OfficeInstalled)) {
                Install-Office21
            }
            # else {
            #     Write-Host -NoNewLine "Press any key to go back to Main Menu "
            #     $x = [System.Console]::ReadKey().KeyChar
            #     Show-OfficeMainMenu    
            #     <# Action when all if and elseif conditions are false #>
            # }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
        'q' {
            Write-Host "Exiting..."
        }
        '0' {
            Show-OfficeMainMenu
        }
        default {
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
    }
}
function Process-OfficeMainMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        'q' {
            Write-Host "Exiting..."
            Stop-Script
        }
        '0' {
            Show-TLMainMenu
            # exit
        }
            '1' {
            Show-OfficeInstallMenu
        }
        '2' {
            Show-OfficeRemoveMenu
        }
        '3' {
            Invoke-Logo
            Write-Host "Running Massgrave.dev Microsoft Activation Scripts" -ForegroundColor Cyan 
            Invoke-MAS
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeMainMenu
        }
        default {
            # Read-Host "Press Enter to continue..."
            # Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeMainMenu
        }
    }
}
function Process-OfficeRemoveMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Running Office Removal Tool with SaRa" -ForegroundColor Cyan
            Invoke-OfficeRemovalTool
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        '2' {
            Invoke-Logo
            Write-Host "Running Office Removal Tool with Office365 Setup" -ForegroundColor Cyan
            Invoke-OfficeRemovalTool -UseSetupRemoval
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        '3' {
            Invoke-Logo
            Write-Host "Running Office Scrubber" -ForegroundColor Cyan
            Install-7ZipIfNeeded
            Invoke-OfficeScrubber
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        'q' {
            Write-Host "Exiting..."
            Stop-Script
        }
        '0' {
            Show-OfficeMainMenu
        }
        default {
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
    }
}
function Process-TLMainMenu-Choice {
    param (
        [string]$choice
    )
  
    switch ($choice) {
        'q' {
            Write-Host "Exiting..."
        }
        '0' {
            Write-Host "Exiting..."
        }
        '1' {
            Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-RestMethod `"$WinUtilUrl`" | Invoke-Expression" -Wait
            Show-TLMainMenu
        }
        '2' {
            # Check if script was run as Administrator, relaunch if not
            if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Clear-Host
                Start-Process -FilePath powershell.exe -ArgumentList "Invoke-RestMethod `"$BinUtilCLIUrl`" | Invoke-Expression" -Wait -NoNewWindow
                Show-TLMainMenu
                }
            else {
                Write-Host "BinUtil can't be run as Administrator..."
                Write-Host "Re-run this command in a non-admin PowerShell window..."
                Write-Host " irm `"$ScriptUrl`" | iex " -ForegroundColor Yellow
                Read-Host "Press Enter to Exit..."
                break
            }

        }
        '3' {
            # Check if script was run as Administrator, relaunch if not
            if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Clear-Host
                Start-Process -FilePath powershell.exe -ArgumentList "Invoke-RestMethod `"$BinUtilGUIUrl`" | Invoke-Expression" -Wait -NoNewWindow
                Show-TLMainMenu
                }
            else {
                Write-Host "BinUtil can't be run as Administrator..."
                Write-Host "Re-run this command in a non-admin PowerShell window..."
                Write-Host " irm `"$ScriptUrl`" | iex " -ForegroundColor Yellow
                Read-Host "Press Enter to Exit..."
                break
            }

        }
        '4' {
            # Check if script was run as Administrator, relaunch if not
            if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Write-Output "OfficeUtil needs to be run as Administrator. Attempting to relaunch."
                Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-RestMethod `"$ScriptUrl`" | Invoke-Expression" 
                break
            }
            Show-OfficeMainMenu
        }
        default {
            # Read-Host "Press Enter to continue..."
            # Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-TLMainMenu
        }
    }
}
  
function Show-OfficeInstallMenu {
  Invoke-Logo
  Write-Host "Install Microsoft Office" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office Deployment Tool"
  Write-Host "2. Install Microsoft Office 365 Business"
  Write-Host "3. Install Microsoft Office 2021 Pro Plus"
  Write-Host "0. Main Office Menu"
  Write-Host "Q. Quit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeInstallMenu-Choice $choice
}
function Show-OfficeMainMenu {
  Invoke-Logo
  Write-Host "Main Menu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office" -ForegroundColor Green
  Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
  Write-Host "0. Main Menu"
  Write-Host "Q. Exit" -ForegroundColor Red

  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeMainMenu-Choice $choice
}
function Show-OfficeRemoveMenu {
  Invoke-Logo
  Write-Host "Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. Run Office Removal Tool with SaRa"
  Write-Host "2. Run Office Removal Tool with Office365 Setup"
  Write-Host "3. Run Office Scrubber"
  Write-Host "0. Main Office Menu"
  Write-Host "Q. Quit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeRemoveMenu-Choice $choice
}
function Show-TLMainMenu {
  Invoke-Logo
  Write-Host "Main Menu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Winutil: Install and Tweak Utility" -ForegroundColor DarkGreen
  Write-Host "2. BinUtil: Recycle Bin Themes (CLI)" -ForegroundColor Magenta
  Write-Host "3. BinUtil: Recycle Bin Themes (GUI)" -ForegroundColor Magenta
  Write-Host "4. OfficeUtil: Install/Remove/Activate Office & Windows" -ForegroundColor Cyan
  Write-Host "Q. Exit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-TLMainMenu-Choice $choice
}
function Stop-Script {
  
  # Clean up: Remove the OfficeUtil folder
  if (Test-Path -Path $OfficeUtilPath -PathType Container) {
    Invoke-Logo
    Write-Host ""
    Write-Host -NoNewLine "Press F to delete $OfficeUtilPath or any other key to quit: "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""
    switch ($choice) {
      'f' {
        Write-Host "Removing "$OfficeUtilPath"\* ..." -ForegroundColor Green
        Remove-Item -LiteralPath $OfficeUtilPath -Force -Recurse
        exit
        }
      'default' {
        exit
      }
    }
    }
  
}

function Test-OfficeInstalled {
  if (Test-Path "C:\Program Files\Microsoft Office") {
    Write-Host "Microsoft Office is already installed." -ForegroundColor Yellow
    Write-Host "Run OfficeRemoverTool and OfficeScrubber to remove the previous installation first."
    Write-Host "Or run Massgrave.dev Microsoft Activation Scripts to activate Office / Windows."

    return $true
  }
  else {
    return $false
  }
}
#===========================================================================
# Shows the form
#===========================================================================

# Invoke-WPFFormVariables

# Toon het hoofdmenu
Show-TLMainMenu


Stop-Script

# Stop-Transcript

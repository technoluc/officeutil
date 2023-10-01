
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

$ScriptUrl = "https://raw.githubusercontent.com/technoluc/scripts/main/officeutil/OfficeUtil.ps1"
$OfficeUtilPath = "C:\OfficeUtil"

$odtInstaller = "C:\OfficeUtil\odtInstaller.exe"
$odtPath = "C:\Program Files\OfficeDeploymentTool"
$setupExe = "C:\Program Files\OfficeDeploymentTool\setup.exe"
$configuration21XML = "C:\Program Files\OfficeDeploymentTool\config21.xml"
$configuration365XML = "C:\Program Files\OfficeDeploymentTool\config365.xml"

# OfficeScrubber
$ArchiveUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberPath = "C:\OfficeUtil\OfficeScrubber"
$ScrubberArchive = "OfficeScrubber_11.7z"
$ScrubberCmd = "OfficeScrubber.cmd"
$ScrubberFullPath = Join-Path -Path $ScrubberPath -ChildPath $ScrubberCmd


$OfficeRemovalToolUrl = "https://raw.githubusercontent.com/technoluc/msoffice-removal-tool/main/msoffice-removal-tool.ps1"
$OfficeRemovalTool = "msoffice-removal-tool.ps1"
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
    [string]$ArchiveUrl,
    [string]$ScrubberPath,
    [string]$ScrubberArchive,
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combineer het pad naar het archief
  $ArchivePath = Join-Path -Path $ScrubberPath -ChildPath $ScrubberArchive

  # Maak de map als deze nog niet bestaat
  if (-not (Test-Path -Path $ScrubberPath -PathType Container)) {
    New-Item -Path $ScrubberPath -ItemType Directory ;
  }

  try {
    # Download het archief
    Invoke-WebRequest -Uri $ArchiveUrl -OutFile $ArchivePath

    # Uitpakken van het archief met het volledige pad naar 7z
    & $7zPath x $ArchivePath -o"$ScrubberPath"

    Write-Host "Het archief is succesvol gedownload en uitgepakt naar: $ScrubberPath"
  }
  catch {
    Write-Host "Er is een fout opgetreden bij het downloaden en uitpakken van het archief: $_"
  }
  finally {
    # Opruimen: Verwijder het gedownloade archief
    Remove-Item -Path $ArchivePath -Force
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
        'q' {
            Write-Host "Afsluiten..."
            exit
        }
        '0' {
            Write-Host "Afsluiten..."
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
            Write-Host " Running Massgrave.dev Microsoft Activation Scripts" -ForegroundColor Cyan 
            Run-MAS
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-MainMenu
        }
        default {
            # Read-Host "Druk op Enter om door te gaan..."
            Write-Host "Ongeldige optie. Probeer opnieuw."
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
            # Voer hier de stappen uit voor Suboptie 1.1
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '2' {
            Invoke-Logo
            Write-Host "Install Microsoft Office 2021 Pro Plus" -ForegroundColor Green
            # Voer hier de stappen uit voor Suboptie 1.2
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '3' {
            Invoke-Logo
            Write-Host "Install Microsoft Office Deployment Tool" -ForegroundColor Green
            # Voer hier de stappen uit voor Suboptie 1.3
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        'q' {
            Write-Host "Afsluiten..."
            exit
        }
        '0' {
            Show-MainMenu
        }
        default {
            # Write-Host "Ongeldige optie. Probeer opnieuw."
            # Read-Host "Druk op Enter om door te gaan..."
            # Read-Host "Druk op Enter om door te gaan..."
            Write-Host "Ongeldige optie. Probeer opnieuw."
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
            # Voer hier de stappen uit voor Suboptie 1.1
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '2' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with Office365 Setup" -ForegroundColor Cyan
            # Voer hier de stappen uit voor Suboptie 1.2
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Run Office Scrubber" -ForegroundColor Cyan
            Get-OfficeScrubber
            Run-OfficeScrubber
            # Voer hier de stappen uit voor Suboptie 1.3
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        'q' {
            Write-Host "Afsluiten..."
            exit
        }
        '0' {
            Show-MainMenu
        }
        default {
            # Write-Host "Ongeldige optie. Probeer opnieuw."
            # Read-Host "Druk op Enter om door te gaan..."
            # Read-Host "Druk op Enter om door te gaan..."
            Write-Host "Ongeldige optie. Probeer opnieuw."
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
  Get-OfficeScrubber -ArchiveUrl $ArchiveUrl -ScrubberPath $ScrubberPath -ScrubberArchive $ScrubberArchive
  Start-Process -Verb runas -FilePath "cmd.exe" -ArgumentList "/C $ScrubberFullPath "
  Read-Host "Press Enter to continue..."
  Remove-Item -Path $ScrubberFullPath -Force
}
function Show-MainMenu {
  Invoke-Logo
  Write-Host "Hoofdmenu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office" -ForegroundColor Green
  Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
  Write-Host "0. Exit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Selecteer een optie (0-3)"
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
  # $choice = Read-Host "Selecteer een optie (0-3)"
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
  # $choice = Read-Host "Selecteer een optie (0-3)"
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

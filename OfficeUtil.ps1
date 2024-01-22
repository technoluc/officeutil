
################################################################################################################
###                                                                                                          ###
### WARNING: This file is automatically generated DO NOT modify this file directly as it will be overwritten ###
###                                                                                                          ###
################################################################################################################

# Start-Transcript $ENV:TEMP\WinScript.log

####################################################################################################
#                                          SET VARIABLES                                           #
####################################################################################################

$OfficeUtilUrl = "https://raw.githubusercontent.com/technoluc/officeutil/main/OfficeUtil.ps1"
$WinUtilUrl = "https://raw.githubusercontent.com/technoluc/winutil/main/winutil.ps1"
$BinUtilUrl = "https://raw.githubusercontent.com/technoluc/recycle-bin-themes/main/RecycleBinThemes.ps1"
$BinUtilGUIUrl = "https://raw.githubusercontent.com/technoluc/recycle-bin-themes/main/RecycleBinThemesGUI.ps1"

$OfficeUtilPath = "C:\OfficeUtil"
$odtPath = "C:\Program Files\OfficeDeploymentTool"

$odtInstallerPath = Join-Path -Path $OfficeUtilPath -ChildPath "odtInstaller.exe"
$setupExePath = Join-Path -Path $odtPath -ChildPath "setup.exe"
$configuration21XMLPath = Join-Path -Path $odtPath -ChildPath "config21.xml"
$configuration365XMLPath = Join-Path -Path $odtPath -ChildPath "config365.xml"
$configuration21XMLUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config21.xml"
$configuration365XMLUrl = "https://github.com/technoluc/winutil/raw/main-custom/office/config365.xml"
$MASUrl = "https://massgrave.dev/get"

# OfficeScrubber
$ScrubberBaseUrl = "https://github.com/abbodi1406/WHD/raw/master/scripts/OfficeScrubber_11.7z"
$ScrubberPath = Join-Path -Path $OfficeUtilPath -ChildPath "OfficeScrubber"
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
function Stop-Script {
  if (Test-Path -Path $OfficeUtilPath -PathType Container) {
    Invoke-Logo
    Write-Host ""
    Write-Host -NoNewLine "Press F to delete $OfficeUtilPath or any other key to quit: "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""

    if ($choice -eq 'f') {
      Write-Host "Removing $OfficeUtilPath\* ..." -ForegroundColor Green
      Remove-Item -LiteralPath $OfficeUtilPath -Force -Recurse
    }
  }
  exit
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
function Install-7ZipIfNeeded {
  $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"

  if (-not $7ZipInstalled) {
      Write-Host "7-Zip is not installed. Installing..."
      $InstallerUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
      $InstallerPath = Join-Path -Path $env:TEMP -ChildPath "7zInstaller.exe"
      
      # Download and install 7-Zip with /S for silent installation
      Download-File -url $InstallerUrl -outputPath $InstallerPath 
      Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
      
      # Check for successful installation
      $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"
      if ($7ZipInstalled) {
          Write-Host "7-Zip has been successfully installed."
      }
      else {
          Write-Host "Error: 7-Zip installation failed."
      }

      # Remove the temporary installation file
      Remove-Item -Path $InstallerPath -Force
  }
  else {
      Write-Host "7-Zip is already installed."
  }
}

function Install-OdtIfNeeded {
  $OdtInstalled = Test-Path "$setupExePath"

  if (-not $OdtInstalled) {
    if (-not (Test-Path "$odtInstallerPath")) {
      
      if (-not (Test-Path "$OfficeUtilPath")) {
        New-Item -Path $OfficeUtilPath -ItemType Directory -Force | Out-Null
        Write-Host "$OfficeUtilPath created"
      }
      if (-not (Test-Path "$odtPath")) {
        New-Item -Path $odtPath -ItemType Directory -Force | Out-Null
        Write-Host "$odtPath created"
      }
      $URL = $(Get-ODTUri)
      Download-File -url $URL -outputPath $odtInstallerPath
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
  }
  else {
    Write-Host ""
    Write-Host "Office Deployment Tool is already installed." -ForegroundColor Green
  }
}

function Install-Office {
  param (
      [string]$product
  )

  Install-OdtIfNeeded

  if ($product -eq "Office 365 Business") {
      $configXMLPath = $configuration365XMLPath
      $downloadURL = $configuration365XMLUrl
      $unattendedArgs = $UnattendedArgs365
  }
  elseif ($product -eq "Office 2021 Pro Plus") {
      $configXMLPath = $configuration21XMLPath
      $downloadURL = $configuration21XMLUrl
      $unattendedArgs = $UnattendedArgs21
  }
  else {
      Write-Host "Invalid product selection."
      return
  }

  if (-not (Test-Path -Path $configXMLPath -PathType Leaf)) {
      Write-Host "Downloading $product Configuration File..." -ForegroundColor Cyan
      Download-File -url $downloadURL -outputPath $configXMLPath
  }

  Write-Host -NoNewline "Install Microsoft '$product'? ( Y / N ): "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""

  switch ($choice) {
      'y' {
          Write-Host "Installation started. Don't close this window" -ForegroundColor Green
          Start-Process -Wait $setupExePath -ArgumentList "$unattendedArgs"
          Write-Host "Installation of Microsoft $product completed." -ForegroundColor Green
          Show-OfficeMainMenu
      }
      'n' {
          Write-Host "Exiting..."
          Show-OfficeInstallMenu
      }
      default {
          Write-Host -NoNewLine "Invalid option. Press any key to try again... "
          $x = [System.Console]::ReadKey().KeyChar
          Invoke-Logo
          Show-OfficeInstallMenu
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

function Process-OfficeInstallMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office Deployment Tool" -ForegroundColor Green
            Install-OdtIfNeeded
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
        '2' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 365 Business" -ForegroundColor Green
            if (-not (Test-OfficeInstalled)) {
                Install-Office -product "Office 365 Business"
            }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
        }
        '3' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 2021 Pro Plus" -ForegroundColor Green
            if (-not (Test-OfficeInstalled)) {
                Install-Office -product "Office 2021 Pro Plus"
            }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeInstallMenu
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
            Show-OfficeInstallMenu
        }
    }
}
function Download-File {
  param (
      [string]$url,
      [string]$outputPath
  )

  try {
      Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
  }
  catch {
      Write-Host "Failed to download $url with error: $_"
  }
}

function Get-ODTUri {
  [CmdletBinding()]
  [OutputType([string])]
  param ()

  $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"

  try {
      $response = Invoke-WebRequest -UseBasicParsing -Uri $url
      $ODTUri = $response.links | Where-Object { $_.outerHTML -like "*click here to download manually*" }
      if ($ODTUri) {
          return $ODTUri.href
      }
      else {
          throw "Failed to retrieve ODT download URL."
      }
  }
  catch {
      throw "Failed to connect to ODT: $url with error $_."
  }
}

function Download-OfficeScrubber{
  Download-File -url $ScrubberBaseUrl -outputPath $ScrubberArchivePath
}


function Extract-OfficeScrubber {
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
    Download-OfficeScrubber
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
function Show-OfficeMainMenu {
    Invoke-Logo
    Write-Host "Main Menu" -ForegroundColor Green
    Write-Host ""
    Write-Host "1. Install Microsoft Office" -ForegroundColor Green
    Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
    Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
    Write-Host "Q. Exit" -ForegroundColor Red

    Write-Host ""
    # $choice = Read-Host "Select an option (0-3)"
    Write-Host -NoNewline "Select option: "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""
    Process-OfficeMainMenu-Choice $choice
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
# Show Main Menu
Show-OfficeMainMenu

Stop-Script

# Stop-Transcript

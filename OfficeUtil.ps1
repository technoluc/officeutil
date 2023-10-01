
################################################################################################################
###                                                                                                          ###
### WARNING: This file is automatically generated DO NOT modify this file directly as it will be overwritten ###
###                                                                                                          ###
################################################################################################################

Start-Transcript $ENV:TEMP\testmenu.log -Append
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
            Read-Host "Druk op Enter om terug te gaan naar het hoofdmenu..."
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

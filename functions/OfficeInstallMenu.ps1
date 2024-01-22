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

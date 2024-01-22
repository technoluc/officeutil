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

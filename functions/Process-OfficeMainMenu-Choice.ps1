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

function Process-OfficeMMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        'q' {
            Write-Host "Exiting..."
            Stop-Script
        }
        '0' {
            Show-TLMainWindow
            # exit
        }
            '1' {
            Show-OfficeSMenu1
        }
        '2' {
            Show-OfficeSMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Running Massgrave.dev Microsoft Activation Scripts" -ForegroundColor Cyan 
            Invoke-MAS
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeMMenu
        }
        default {
            # Read-Host "Press Enter to continue..."
            # Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeMMenu
        }
    }
}

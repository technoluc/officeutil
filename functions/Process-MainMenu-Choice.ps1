function Process-MainMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        'q' {
            Write-Host "Exiting..."
            exit
        }
        '0' {
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

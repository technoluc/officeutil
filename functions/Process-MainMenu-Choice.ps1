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

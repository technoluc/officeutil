function Process-MainMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
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
            Write-Host "Ongeldige optie. Probeer opnieuw."
            Read-Host "Druk op Enter om door te gaan..."
            Show-MainMenu
        }
    }
}

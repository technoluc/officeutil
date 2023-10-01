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

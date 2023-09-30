function Process-SubMenu2-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with SaRa"
            # Voer hier de stappen uit voor Suboptie 1.1
            Read-Host "Druk op Enter om verder te gaan..."
            Show-SubMenu2
        }
        '2' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with Office365 Setup"
            # Voer hier de stappen uit voor Suboptie 1.2
            Read-Host "Druk op Enter om verder te gaan..."
            Show-SubMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Run Office Scrubber"
            # Voer hier de stappen uit voor Suboptie 1.3
            Read-Host "Druk op Enter om verder te gaan..."
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
            Write-Host "Ongeldige optie. Probeer opnieuw."
            Read-Host "Druk op Enter om door te gaan..."
            Show-SubMenu2
        }
    }
}

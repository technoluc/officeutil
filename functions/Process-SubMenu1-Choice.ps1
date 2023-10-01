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

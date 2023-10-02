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

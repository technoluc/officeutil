function Process-SubMenu1-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office Deployment Tool" -ForegroundColor Green
            # Perform the steps for Suboption 1.1 here
            Get-OdtIfNeeded
            # Perform the steps for Suboption 1.1 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '2' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 365 Business" -ForegroundColor Green
            # Perform the steps for Suboption 1.2 here
            if (-not (Test-OfficeInstalled)) {
                Install-Office365
            }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        '3' {
            Invoke-Logo
            Write-Host "Installing Microsoft Office 2021 Pro Plus" -ForegroundColor Green
            if (-not (Test-OfficeInstalled)) {
                Install-Office21
            }
            # else {
            #     Write-Host -NoNewLine "Press any key to go back to Main Menu "
            #     $x = [System.Console]::ReadKey().KeyChar
            #     Show-MainMenu    
            #     <# Action when all if and elseif conditions are false #>
            # }
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
        'q' {
            Write-Host "Exiting..."
        }
        '0' {
            Show-MainMenu
        }
        default {
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu1
        }
    }
}

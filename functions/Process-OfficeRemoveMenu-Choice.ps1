function Process-OfficeRemoveMenu-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Running Office Removal Tool with SaRa" -ForegroundColor Cyan
            Invoke-OfficeRemovalTool
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        '2' {
            Invoke-Logo
            Write-Host "Running Office Removal Tool with Office365 Setup" -ForegroundColor Cyan
            Invoke-OfficeRemovalTool -UseSetupRemoval
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        '3' {
            Invoke-Logo
            Write-Host "Running Office Scrubber" -ForegroundColor Cyan
            Install-7ZipIfNeeded
            Invoke-OfficeScrubber
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
        'q' {
            Write-Host "Exiting..."
            Stop-Script
        }
        '0' {
            Show-OfficeMainMenu
        }
        default {
            Write-Host -NoNewLine "Invalid option. Press any key to try again... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-OfficeRemoveMenu
        }
    }
}

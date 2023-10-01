function Process-SubMenu2-Choice {
    param (
        [string]$choice
    )

    switch ($choice) {
        '1' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with SaRa" -ForegroundColor Cyan
            # Perform the steps for Suboption 1.1 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '2' {
            Invoke-Logo
            Write-Host "Run Office Removal Tool with Office365 Setup" -ForegroundColor Cyan
            # Perform the steps for Suboption 1.2 here
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        '3' {
            Invoke-Logo
            Write-Host "Run Office Scrubber" -ForegroundColor Cyan
            Get-7ZipIfNeeded
            Invoke-OfficeScrubber
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
        'q' {
            Write-Host "Exiting..."
            #exit
        }
        '0' {
            Show-MainMenu
        }
        default {
            Write-Host "Invalid option. Please try again."
            Write-Host -NoNewLine "Press any key to continue... "
            $x = [System.Console]::ReadKey().KeyChar
            Show-SubMenu2
        }
    }
}

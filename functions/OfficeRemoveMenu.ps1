function Show-OfficeRemoveMenu {
  Invoke-Logo
  Write-Host "Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. Run Office Removal Tool with SaRa"
  Write-Host "2. Run Office Removal Tool with Office365 Setup"
  Write-Host "3. Run Office Scrubber"
  Write-Host "0. Main Office Menu"
  Write-Host "Q. Quit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeRemoveMenu-Choice $choice
}

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

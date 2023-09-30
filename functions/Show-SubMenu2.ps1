function Show-SubMenu2 {
  Invoke-Logo
  Write-Host "Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. Run Office Removal Tool with SaRa"
  Write-Host "2. Run Office Removal Tool with Office365 Setup"
  Write-Host "3. Run Office Scrubber"
  Write-Host "0. Terug naar hoofdmenu"
  Write-Host "Q. Quit"
  Write-Host ""
  $choice = Read-Host "Selecteer een optie (0-3)"
  Process-SubMenu2-Choice $choice
}

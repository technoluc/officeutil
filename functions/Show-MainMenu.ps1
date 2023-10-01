function Show-MainMenu {
  Invoke-Logo
  Write-Host "Hoofdmenu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office" -ForegroundColor Green
  Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
  Write-Host "0. Exit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Selecteer een optie (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-MainMenu-Choice $choice
}

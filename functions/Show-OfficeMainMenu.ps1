function Show-OfficeMainMenu {
  Invoke-Logo
  Write-Host "Main Menu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office" -ForegroundColor Green
  Write-Host "2. Uninstall Microsoft Office" -ForegroundColor Yellow
  Write-Host "3. Activate Microsoft Office / Windows" -ForegroundColor Cyan
  Write-Host "0. Main Menu"
  Write-Host "Q. Exit" -ForegroundColor Red

  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeMainMenu-Choice $choice
}

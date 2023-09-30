function Show-MainMenu {
  Invoke-Logo
  Write-Host "Hoofdmenu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. $MainMenuItem1"
  Write-Host "2. $MainMenuItem2"
  Write-Host "3. $MainMenuItem3"
  Write-Host "0. Exit" -ForegroundColor Red
  Write-Host ""
  $choice = Read-Host "Selecteer een optie (0-3)"
  Process-MainMenu-Choice $choice
}

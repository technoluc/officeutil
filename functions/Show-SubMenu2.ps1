function Show-SubMenu2 {
  Clear-Host
  Invoke-Logo
  Write-Host "$MainMenuItem2" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. $SubMenu2Item1"
  Write-Host "2. $SubMenu2Item2"
  Write-Host "3. $SubMenu2Item3"
  Write-Host "0. Terug naar hoofdmenu"
  Write-Host ""
  $choice = Read-Host "Selecteer een optie (0-3)"
  Process-SubMenu2-Choice $choice
}

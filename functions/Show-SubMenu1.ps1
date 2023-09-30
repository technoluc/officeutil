function Show-SubMenu1 {
  Clear-Host
  Invoke-Logo
  Write-Host "$MainMenuItem1" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "1. $SubMenu1Item1"
  Write-Host "2. $SubMenu1Item2"
  Write-Host "3. $SubMenu1Item3"
  Write-Host "0. Terug naar hoofdmenu"
  Write-Host ""
  $choice = Read-Host "Selecteer een optie (0-3)"
  Process-SubMenu1-Choice $choice
}

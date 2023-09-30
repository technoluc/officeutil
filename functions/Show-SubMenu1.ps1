function Show-SubMenu1 {
  Invoke-Logo
  Write-Host "Install Microsoft Office" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office 365 Business"
  Write-Host "2. Install Microsoft Office 2021 Pro Plus"
  Write-Host "3. Install Microsoft Office Deployment Tool"
  Write-Host "0. Terug naar hoofdmenu"
  Write-Host "Q. Quit"
  Write-Host ""
  $choice = Read-Host "Selecteer een optie (0-3)"
  Process-SubMenu1-Choice $choice
}

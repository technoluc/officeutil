function Show-SubMenu1 {
  Invoke-Logo
  Write-Host "Install Microsoft Office" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office 365 Business"
  Write-Host "2. Install Microsoft Office 2021 Pro Plus"
  Write-Host "3. Install Microsoft Office Deployment Tool"
  Write-Host "0. Main menu"
  Write-Host "Q. Quit"
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-SubMenu1-Choice $choice
}

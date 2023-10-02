function Show-OfficeInstallMenu {
  Invoke-Logo
  Write-Host "Install Microsoft Office" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Install Microsoft Office Deployment Tool"
  Write-Host "2. Install Microsoft Office 365 Business"
  Write-Host "3. Install Microsoft Office 2021 Pro Plus"
  Write-Host "0. Main Office Menu"
  Write-Host "Q. Quit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-OfficeInstallMenu-Choice $choice
}

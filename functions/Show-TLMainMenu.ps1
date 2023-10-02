function Show-TLMainMenu {
  Invoke-Logo
  Write-Host "Main Menu" -ForegroundColor Green
  Write-Host ""
  Write-Host "1. Winutil: Install and Tweak Utility" -ForegroundColor DarkGreen
  Write-Host "2. BinUtil: Recycle Bin Themes" -ForegroundColor Magenta
  Write-Host "3. OfficeUtil: Install/Remove/Activate Office & Windows" -ForegroundColor Cyan
  Write-Host "Q. Exit" -ForegroundColor Red
  Write-Host ""
  # $choice = Read-Host "Select an option (0-3)"
  Write-Host -NoNewline "Select option: "
  $choice = [System.Console]::ReadKey().KeyChar
  Write-Host ""
  Process-TLMainMenu-Choice $choice
}

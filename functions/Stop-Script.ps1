function Stop-Script {
  
  # Clean up: Remove the downloaded archive
  if (Test-Path -Path $OfficeUtilPath -PathType Container) {
    Write-Host "Removing "$OfficeUtilPath"\* ..." -ForegroundColor Green
    Remove-Item -LiteralPath $OfficeUtilPath -Force -Recurse
  }
  Write-Host "Exiting... "

}
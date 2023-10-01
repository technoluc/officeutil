function Stop-Script {
  # Clean up: Remove the downloaded archive
  if (Test-Path -Path $ScrubberPath -PathType Container) {
    Write-Host "Removing "$ScrubberPath"\* ..." -ForegroundColor Green
    Remove-Item -LiteralPath $ScrubberPath -Force -Recurse
  }

}
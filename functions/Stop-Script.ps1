function Stop-Script {
  if (Test-Path -Path $OfficeUtilPath -PathType Container) {
    Invoke-Logo
    Write-Host ""
    Write-Host -NoNewLine "Press F to delete $OfficeUtilPath or any other key to quit: "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""

    if ($choice -eq 'f') {
      Write-Host "Removing $OfficeUtilPath\* ..." -ForegroundColor Green
      Remove-Item -LiteralPath $OfficeUtilPath -Force -Recurse
    }
  }
  exit
}

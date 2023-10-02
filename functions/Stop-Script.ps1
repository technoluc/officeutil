function Stop-Script {
  
  # Clean up: Remove the OfficeUtil folder
  if (Test-Path -Path $OfficeUtilPath -PathType Container) {
    Invoke-Logo
    Write-Host ""
    Write-Host -NoNewLine "Press F to delete $OfficeUtilPath or any other key to quit: "
    $choice = [System.Console]::ReadKey().KeyChar
    Write-Host ""
    switch ($choice) {
      'f' {
        Write-Host "Removing "$OfficeUtilPath"\* ..." -ForegroundColor Green
        Remove-Item -LiteralPath $OfficeUtilPath -Force -Recurse
        exit
        }
      'default' {
        exit
      }
    }
    }
  
}


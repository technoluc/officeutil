function Install-7ZipIfNeeded {
  $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"

  if (-not $7ZipInstalled) {
      Write-Host "7-Zip is niet geïnstalleerd. Installeren..."
      $InstallerUrl = "https://www.7-zip.org/a/7z2301-x64.exe"
      $InstallerPath = Join-Path -Path $env:TEMP -ChildPath "7zInstaller.exe"
      
      # Download het 7-Zip-installatieprogramma
      Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerPath -UseBasicParsing
      
      # Installeer 7-Zip met /S om stil te installeren
      Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait
      
      # Controleren op succesvolle installatie
      $7ZipInstalled = Test-Path "C:\Program Files\7-Zip\7z.exe"
      if ($7ZipInstalled) {
          Write-Host "7-Zip is succesvol geïnstalleerd."
      } else {
          Write-Host "Fout: 7-Zip-installatie is mislukt."
      }

      # Verwijder het tijdelijke installatiebestand
      Remove-Item -Path $InstallerPath -Force
  } else {
  }
}

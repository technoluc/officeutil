function Get-OfficeScrubber {
  param (
    [string]$ScrubberBaseUrl,
    [string]$OfficeUtilPath,
    [string]$ScrubberArchiveName,
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combineer het pad naar het archief
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Maak de map als deze nog niet bestaat
  if (-not (Test-Path -Path $OfficeUtilPath -PathType Container)) {
    New-Item -Path $OfficeUtilPath -ItemType Directory ;
  }

  try {
    # Download het archief
    Invoke-WebRequest -Uri $ScrubberBaseUrl -OutFile $ScrubberArchivePath

    # Uitpakken van het archief met het volledige pad naar 7z
    & $7zPath x $ScrubberArchivePath -o"$OfficeUtilPath"

    Write-Host "Het archief is succesvol gedownload en uitgepakt naar: $OfficeUtilPath"
  }
  catch {
    Write-Host "Er is een fout opgetreden bij het downloaden en uitpakken van het archief: $_"
  }
  finally {
    # Opruimen: Verwijder het gedownloade archief
    Remove-Item -Path $ScrubberArchivePath -Force
  }
}

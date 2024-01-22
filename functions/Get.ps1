function Download-File {
  param (
      [string]$url,
      [string]$outputPath
  )

  try {
      Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
  }
  catch {
      Write-Host "Failed to download $url with error: $_"
  }
}

function Get-ODTUri {
  [CmdletBinding()]
  [OutputType([string])]
  param ()

  $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"

  try {
      $response = Invoke-WebRequest -UseBasicParsing -Uri $url
      $ODTUri = $response.links | Where-Object { $_.outerHTML -like "*click here to download manually*" }
      if ($ODTUri) {
          return $ODTUri.href
      }
      else {
          throw "Failed to retrieve ODT download URL."
      }
  }
  catch {
      throw "Failed to connect to ODT: $url with error $_."
  }
}

function Download-OfficeScrubber{
  Download-File -url $ScrubberBaseUrl -outputPath $ScrubberArchivePath
}


function Extract-OfficeScrubber {
  param (
    [string]$7zPath = "C:\Program Files\7-Zip\7z.exe"
  )

  # Combine the path to the archive
  $ScrubberArchivePath = Join-Path -Path $OfficeUtilPath -ChildPath $ScrubberArchiveName

  # Create the directory if it doesn't exist yet
  if (-not (Test-Path -Path $ScrubberPath -PathType Container)) {
    New-Item -Path $ScrubberPath -ItemType Directory -Force | Out-Null
  }

  try {
    # Download and extract the archive using 7-Zip
    Download-OfficeScrubber
    & $7zPath x $ScrubberArchivePath -o"$ScrubberPath" -y > $null

    Write-Host "The archive has been successfully downloaded and extracted to: $ScrubberPath" -ForegroundColor Green
  }
  catch {
    Write-Host "An error occurred while downloading and extracting the archive: $_" -ForegroundColor Red
  }
  finally {
    # Clean up: Remove the downloaded archive
    if (Test-Path -Path $ScrubberArchivePath -PathType Leaf) {
      Remove-Item -Path $ScrubberArchivePath -Force
    }
  }
}


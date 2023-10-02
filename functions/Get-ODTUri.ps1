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
      } else {
          throw "Failed to retrieve ODT download URL."
      }
  } catch {
      throw "Failed to connect to ODT: $url with error $_."
  }
}

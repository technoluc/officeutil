function Invoke-MAS {
  # Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "Invoke-WebRequest -useb https://massgrave.dev/get | Invoke-Expression" -Wait
  Invoke-RestMethod https://massgrave.dev/get | Invoke-Expression
}

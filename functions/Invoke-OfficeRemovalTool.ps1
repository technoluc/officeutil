function Invoke-OfficeRemovalTool {
    param (
        [switch]$UseSetupRemoval
    )

    if (-not (Test-Path -Path $OfficeUtilPath -PathType Container)) {
        New-Item -Path $OfficeUtilPath -ItemType Directory | Out-Null
    }

    if ($UseSetupRemoval.IsPresent) {
        $Command = "powershell -ExecutionPolicy Bypass -File $OfficeRemovalToolPath -SuppressReboot -UseSetupRemoval"
    }
    else {
        $Command = "powershell -ExecutionPolicy Bypass -File $OfficeRemovalToolPath -SuppressReboot"
    }

    Invoke-WebRequest -Uri $OfficeRemovalToolUrl -OutFile $OfficeRemovalToolPath
    Invoke-Expression $Command
}

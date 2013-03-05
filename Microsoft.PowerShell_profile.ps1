$PSProfileFolder = Split-Path -Path $PROFILE
$PSHistoryPath = "$PSProfileFolder\History.csv"
$PSMaximumHistoryCount = 1KB

if (!(Test-Path $PSProfileFolder -PathType Container)) {
	New-Item $PSProfileFolder -ItemType Directory -Force
}

if (Test-path $PSHistoryPath) {
	Import-CSV $PSHistoryPath | Add-History
}

Register-EngineEvent PowerShell.Exiting –Action {
	Get-History -Count $PSMaximumHistoryCount | Export-CSV $PSHistoryPath
    [enviroment]::Exit(0)
} | out-null
#region History
$PSProfileFolder = Split-Path -Path $PROFILE
$PSHistoryPath = "$PSProfileFolder\History.csv"
$PSMaximumHistoryCount = 1KB

if (!(Test-Path $PSProfileFolder -PathType Container)) {
	New-Item $PSProfileFolder -ItemType Directory -Force
}

if (Test-path $PSHistoryPath) {
	Import-CSV $PSHistoryPath | Add-History
}

Register-EngineEvent PowerShell.Exiting -Action {
	Get-History -Count $PSMaximumHistoryCount | Export-CSV $PSHistoryPath
    [enviroment]::Exit(0)
} | out-null
#endregion History

function prompt {
	Write-Host ('{') -NoNewline 
	Write-Host ('{0} ' -f ((Get-History -Count 1).Id + 1)) -NoNewLine -ForegroundColor Red
	Write-Host ('{0}'  -f (get-date -Format "hh:mm")) -NoNewLine -Fore Cyan
	Write-Host (' {0}' -f $pwd) -ForegroundColor Gray -NoNewline
	Write-Host ('}') 
}

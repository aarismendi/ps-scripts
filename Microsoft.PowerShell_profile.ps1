#region Alias Helpers
New-Alias which get-command
#endregion

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
	Write-Host ('{0}'   -f (get-date -Format "hh:mm")) -NoNewLine -Fore Cyan
	Write-Host (' {0}' -f $pwd) -ForegroundColor Gray -NoNewline
	Write-Host ('}') 
}

function Unblock-File {
	[cmdletbinding()]
	param (
		[parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[string] $FilePath
	)
	begin {
		#http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
		Add-Type -Namespace PsFile -Name NtfsSecurity -MemberDefinition @"
			[DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool DeleteFile(string name);
			public static bool Unblock(string filePath) {
				return DeleteFile(filePath + ":Zone.Identifier");
			}
"@
	}
	process {
		try {
			#Discard the boolean result.
			if ([io.file]::exists($FilePath)) {
				Write-Verbose "Unblocking $FilePath"
				[PsFile.NtfsSecurity]::Unblock($FilePath) > $null
			} else {
				Write-Verbose "Ignoring non-file $FilePath"
			}
		} catch {Write-Error (
				"Failed to unblock file '{0}'. The error was: '{1}'." -f $FilePath, $_)}
	}
	end {}
}

function which ($command) {
    Get-Command -Name $command -ErrorAction SilentlyContinue | 
        Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
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
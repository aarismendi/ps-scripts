function Unblock-File {
	[cmdletbinding()]
	param ([parameter(Mandatory=$true, ValueFromPipeline=$true)] [string] $FilePath)
	begin {
		Add-Type -Namespace PsFile -Name NtfsSecurity -MemberDefinition @"
			[DllImport("Kernel32.dll")] 
            public static extern int GetLastError();
            public static int GetLastWin32Error(string filePath) {
                return GetLastError();
            }
            // http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
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
				if (-not [PsFile.NtfsSecurity]::Unblock($FilePath)) {
                    $last_error = [PsFile.NtfsSecurity]::GetLastWin32Error()
                    Write-Error "Failed to unblock '$FilePath' the error code is '$last_error'."
                }
			} else { Write-Verbose "Ignoring non-file $FilePath" }
		} catch { Write-Error ("Failed to unblock file '{0}'. The error was: '{1}'." -f $FilePath, $_) } 
    }
}
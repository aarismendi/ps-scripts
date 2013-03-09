function Test-IsAdmin {
	try {
		$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
		$principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
		return $principal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
	} catch {
		throw "Failed to determine if the current user has elevated privileges. The error was: '$_'." -f $_
	}
}

function Unblock-File {
	[cmdletbinding()]
	param ([parameter(Mandatory=$true,  ValueFromPipelineByPropertyName=$true)]
			[alias("FullName")] [string] $FilePath
	)
	begin {
		Add-Type -Namespace PsIO -Name File -MemberDefinition @"
			public static int GetLastWin32Error() {
				return Marshal.GetLastWin32Error();
			}
            // http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
            [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool DeleteFile(string name);
            public static bool Win32DeleteFile(string filePath) {
				return DeleteFile(filePath);
			}
"@
	}
	process {
		try {
			if (Test-Path -Path $FilePath -PathType Leaf) {
				Write-Verbose "Unblocking $FilePath"
				$zone_id_path = $FilePath + ':Zone.Identifier'
				$Win32_ERROR_FILE_NOT_FOUND = 2
				if (-not [PsIO.File]::Win32DeleteFile($zone_id_path)) {
                    $last_error = [PsIO.File]::GetLastWin32Error()
					if ($last_error -ne $Win32_ERROR_FILE_NOT_FOUND) {
						Write-Error "Failed to unblock '$FilePath' the Win32 return code is '$last_error'."
					}
                }
			} else { Write-Verbose "Ignoring non-file $FilePath" }
		} catch { Write-Error ("Failed to unblock file '{0}'. The error was: '{1}'." -f $FilePath, $_) } 
    }
}
function Test-IsAdmin {  
	return (([Security.Principal.WindowsPrincipal] `
				[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
					[Security.Principal.WindowsBuiltInRole] "Administrator"))	
}

function Unblock-File {
	[cmdletbinding()]
	param ([parameter(Mandatory=$true,  ValueFromPipelineByPropertyName=$true)]
			[alias("FullName")] [string] $FilePath
	)
	begin {
		Add-Type -Namespace Win32 -Name pinvoke -MemberDefinition @"
			// http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
			[DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
			[return: MarshalAs(UnmanagedType.Bool)]
			private static extern bool DeleteFile(string name);
			public static int Win32DeleteFile(string filePath) {
				bool is_deleted = DeleteFile(filePath);
				return Marshal.GetLastWin32Error();
			}
"@
	}
	process {
		try {
			$FilePath = (Resolve-Path -Path $FilePath -ErrorAction Stop).Path
			if (Test-Path -Path $FilePath -PathType Leaf) {
				Write-Verbose "Unblocking $FilePath"
				$Win32_SUCCESS = 0
				$Win32_FILE_NOT_FOUND = 2
				$zone_id_path = $FilePath + ':Zone.Identifier'
				$result_code = [Win32.pinvoke]::Win32DeleteFile($zone_id_path)
				if (-not ($Win32_SUCCESS,$Win32_FILE_NOT_FOUND) -contains $result_code) {
					Write-Error "Failed to unblock '$FilePath' the Win32 return code is '$result_code'."
				}
			} else { Write-Verbose "Ignoring non-file $FilePath" }
		} catch { Write-Error ("Failed to unblock file. The error was: '{0}'." -f $_) } 
	}
}
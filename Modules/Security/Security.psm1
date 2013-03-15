function Test-IsAdmin {  
	return (([Security.Principal.WindowsPrincipal] `
				[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
					[Security.Principal.WindowsBuiltInRole] "Administrator"))	
}

function Unblock-File {
	[cmdletbinding(DefaultParameterSetName="ByName", SupportsShouldProcess=$True)]
	param (
        [parameter(Mandatory=$true, ParameterSetName="ByName", Position=0)] [string] $FilePath,
        [parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true)] $InputObject
	)
	begin {
		Add-Type -Namespace Win32 -Name PInvoke -MemberDefinition @"
    // http://msdn.microsoft.com/en-us/library/windows/desktop/aa363915(v=vs.85).aspx
    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteFile(string name);
    public static int Win32DeleteFile(string filePath) {bool gone = DeleteFile(filePath); return Marshal.GetLastWin32Error();}

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    static extern int GetFileAttributes(string lpFileName);
    public static bool Win32FileExists(string filePath) {return GetFileAttributes(filePath) != -1;}
"@
	}
	process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName' { 
                    $input_path = (Resolve-Path -Path $FilePath).Path 
                    if ([IO.File]::Exists($input_path)) { $input_file = Get-Item -Path $input_path }
                }
            'ByInput' { if ($InputObject -is [System.IO.FileInfo]) { $input_file = $InputObject } }
        }
        if ($input_file) {     
            if ([Win32.PInvoke]::Win32FileExists($input_file.FullName + ':Zone.Identifier')) {
                if ($PSCmdlet.ShouldProcess($input_file.FullName)) {
                    $result_code = [Win32.PInvoke]::Win32DeleteFile($input_file.FullName + ':Zone.Identifier')
                    if ($result_code -ne 0) {
                        Write-Error ("Failed to unblock '{0}' the Win32 return code is '{1}'." -f $input_file.FullName, $result_code)
                    }
                }
            }
		}
	}
}
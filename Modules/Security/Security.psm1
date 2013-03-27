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
        public static int Win32DeleteFile(string filePath) {
            bool is_gone = DeleteFile(filePath); return Marshal.GetLastWin32Error();}

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetFileAttributes(string lpFileName);
        public static bool Win32FileExists(string filePath) {return GetFileAttributes(filePath) != -1;}
"@
    }
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByName'  {$input_paths = Resolve-Path -Path $FilePath | ? {[IO.File]::Exists($_.Path)} | Select -Exp Path}
            'ByInput' {if ($InputObject -is [System.IO.FileInfo]) {$input_paths = $InputObject.FullName}}
        }
        $input_paths | % {     
            if ([Win32.PInvoke]::Win32FileExists($_ + ':Zone.Identifier')) {
                if ($PSCmdlet.ShouldProcess($_)) {
                    $result_code = [Win32.PInvoke]::Win32DeleteFile($_ + ':Zone.Identifier')
                    if ([Win32.PInvoke]::Win32FileExists($_ + ':Zone.Identifier')) {
                        Write-Error ("Failed to unblock '{0}' the Win32 return code is '{1}'." -f $_, $result_code)
                    }
                }
            }
        }
    }
}

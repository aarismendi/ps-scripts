$ErrorActionPreference = "Stop"
$this_path = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
$profile_dir = Split-Path $PROFILE
$profile_files = @(
	'Microsoft.PowerShellISE_profile.ps1'
	'Microsoft.PowerShell_profile.ps1'
	'profile.ps1'
)
if (-not [IO.Directory]::Exists($profile_dir)) {
	New-Item -Path $profile_dir -ItemType Directory -Force | Out-Null
}
push-location $this_path
foreach ($file in $profile_files) {
	$src_file = (Get-Item $file).FullName
	copy-item $src_file $profile_dir -Force
}
pop-location

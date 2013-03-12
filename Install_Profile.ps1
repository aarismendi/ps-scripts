$ErrorActionPreference = 'Stop'
$this_path = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
$sec_mod_path = Join-Path -Path $this_path -ChildPath 'Modules\Security\Security.psm1' | Get-Item | Select -ExpandProperty FullName
Import-Module -Name $sec_mod_path -Force

if (-not (Test-IsAdmin)) {
	#Re-launch elevated if not already running elevated.
	$powershellArgs = '-noprofile -nologo -executionpolicy bypass -file "{0}"' -f $MyInvocation.MyCommand.Path
	Start-Process -FilePath 'powershell.exe' -ArgumentList $powershellArgs -Verb RunAs
	exit 
}

#region Install profiles.
$profile_dir = Split-Path $PROFILE
if (-not (Test-Path -Path $profile_dir -PathType Container)) {
    New-Item -Path $profile_dir -ItemType Directory -Force
}
$profile_files = @(
	'Microsoft.PowerShellISE_profile.ps1'
	'Microsoft.PowerShell_profile.ps1'
	'profile.ps1'
)
push-location $this_path
foreach ($file in $profile_files) {
	$src_file = (Get-Item $file).FullName
	copy-item $src_file $profile_dir -Force
    Unblock-File $src_file
}
pop-location
#endregion

#region Install Fonts.
Add-Type -AssemblyName System.Drawing | Out-Null
$fonts = New-Object System.Drawing.Text.InstalledFontCollection
$font_fams = $fonts.Families
$console_font_reg_path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Console\TrueTypeFont'
$console_font_id = Get-ItemProperty $console_font_reg_path | % {
        $_.psbase.properties | ? {$_.Name.StartsWith('0')} 
    } | Select -exp Name -Last 1
$font_path = (Get-Item (Join-Path $this_path 'fonts')).FullName
$FONTS = 0x14
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace($FONTS)
dir -Path $font_path -Filter *.ttf | % {
    $font_col = New-Object System.Drawing.Text.PrivateFontCollection
    $font_col.AddFontFile($_.FullName)
    $font_name = $font_col.Families[0].Name
    
    if (-not $font_fams -contains $font_name) {
        $objFolder.CopyHere($_.FullName)
    }

    $console_font_id += '0'
    $is_registered = (Get-ItemProperty $console_font_reg_path).psbase.properties | ? {$_.Value -eq $font_name}    
    if (-not $is_registered) {
        New-ItemProperty -Path $console_font_reg_path -Name $console_font_id -Value $font_name
    }
}
#endregion
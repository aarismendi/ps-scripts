cls
$ErrorActionPreference = 'Stop'

$path = "~\Documents\GitHub\ps-scripts\Modules\WinFirewall\WinFirewall.psm1"
Import-Module -Name $path -Force -Verbose

$rule = New-FirewallRule -Name all1 -Description 'all' -Port 80 -Protocol tcp -Action allow -Direction in -Profile all
Read-Host 'new rule created'

#Set-FirewallRule

Remove-FirewallRule $rule.Name
Read-Host 'rule removed'

Set-FirewallProfile -DisableUnicastResponse False -Name private
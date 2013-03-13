cls
$ErrorActionPreference = 'Stop'
#$_fw = New-Object -ComObject HNetCfg.FwPolicy2

Import-Module -Name "~\Documents\GitHub\ps-scripts\Modules\WinFirewall\WinFirewall.psm1" -Force -Verbose

Remove-FirewallRule -Name *

$rule = New-FirewallRule -Name foo1 -Description bar1 -Port 80 -Protocol tcp -Action allow -Direction in -Profile all
$rule2 = New-FirewallRule -Name foo1 -Description bar2 -Port 81 -Protocol tcp -Action allow -Direction in -Profile all

Get-FirewallRule | Remove-FirewallRule -verbose 

<#
$rule = New-FirewallRule -Name $rule_name -Port 80 -Protocol tcp -Action allow -Direction in -Profile all
Read-Host 'new rule created'

$rule | Set-FirewallRule -NewName bob -Enabled False -Verbose

Remove-FirewallRule $rule.Name
Read-Host 'rule removed'
#>
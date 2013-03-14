cls
$ErrorActionPreference = 'Stop'
#$_fw = New-Object -ComObject HNetCfg.FwPolicy2

Import-Module -Name "~\Documents\GitHub\ps-scripts\Modules\WinFirewall\WinFirewall.psm1" -Force -Verbose

Remove-FirewallRule -All -Verbose

$rule = New-FirewallRule -Name foo1 -Description '|bar1' -Port 80 -Protocol tcp -Action allow -Direction in -Profile all -Verbose
$rule2 = New-FirewallRule -Verbose -Name foo2 -Description bar2 -Port 81 -Protocol tcp -Action allow -Direction in -Profile all `
            -RemoteAddress '2001:0:4137:9e76:4cd:2c0f:bb9b:617b','192.168.9.99/16','192.168.9.100-192.168.9.150','192.168.9.99',localsubnet,dns,dhcp,wins,defaultgateway 

#Get-FirewallRule | Remove-FirewallRule -verbose

$rule | Set-FirewallRule -LocalPort 90 -RemotePort 1000 -Direction out -Verbose

#Remove-FirewallRule -RemotePort 1001 -Verbose

<#
$rule = New-FirewallRule -Name $rule_name -Port 80 -Protocol tcp -Action allow -Direction in -Profile all
Read-Host 'new rule created'

$rule | Set-FirewallRule -NewName bob -Enabled False -Verbose

Remove-FirewallRule $rule.Name
Read-Host 'rule removed'
#>
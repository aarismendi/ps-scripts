cls

Import-Module C:\Users\andy\Documents\GitHub\ps-scripts\WinFirewall.psm1 -Force -Verbose

$rule = New-FirewallRule -Name all1 -Description 'all' -Ports 80 -Protocol tcp -Action allow -Direction in -Profile all
Read-Host 'new rule created'

Disable-FirewallRule $rule.Name | Out-Null
Read-Host 'rule disabled'

Enable-FirewallRule $rule.Name 
Read-Host 'rule enabled'

Remove-FirewallRule $rule.Name
Read-Host 'rule removed'

Set-InboundTrafficAction -Profile private,domain,public -Action block
Read-Host 'inbound traffic set'

Set-OutboundTrafficAction -Profile private,domain,public -Action default
Read-Host 'outbound traffic set'
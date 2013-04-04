#region
$test = 'Get-FirewallRule: '
Remove-FirewallRule -All

New-FirewallRule -Name foo | Out-Null
New-FirewallRule -Name foo2 | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 2) {$r; throw "Failed $test"}
Remove-FirewallRule -All
#region
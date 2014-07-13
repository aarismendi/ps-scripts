#region New-FirewallRule
$test = 'Get-FirewallProfile: '

Get-fire
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test name(foo)"}
Remove-FirewallRule -All
#endregion
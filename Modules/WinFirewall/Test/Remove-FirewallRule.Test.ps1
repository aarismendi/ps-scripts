$test = 'Remove-FirewallRule: '

New-FirewallRule -Name foo | Out-Null
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test name(foo)"}
Remove-FirewallRule -All
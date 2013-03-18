$test = 'Remove-FirewallRule: '

Remove-FirewallRule -All

New-FirewallRule -Name foo | Out-Null
New-FirewallRule -Name foo2 | Out-Null
Remove-FirewallRule -Name foo -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Inbound | Out-Null
New-FirewallRule -Name foo -Direction Outbound | Out-Null
Remove-FirewallRule -Name foo -Direction Outbound -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Direction(Outbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain | Out-Null
New-FirewallRule -Name foo -Profile Public | Out-Null
Remove-FirewallRule -Name foo -Profile Domain -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Profile(Domain)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress Defaultgateway | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(LocalSubnet,Defaultgateway)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120-192.168.9.130)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120/16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Out-Null
New-FirewallRule -Name foo -Protocol udp | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 40 | Out-Null
New-FirewallRule -Name foo -Protocol 50 | Out-Null
Remove-FirewallRule -Name foo -Protocol 50 -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(50)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
Remove-FirewallRule -Name foo -Protocol 'ICMPv4:3,5' -Verbose
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(ICMPv4:3,5)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
Remove-FirewallRule -Name foo -Protocol any -Verbose -Debug
$r = @(Get-FirewallRule)
if ($r -ne $null) {$r; throw "Failed $test name(foo),Protocol(any)"}
Remove-FirewallRule -All

<#
-RemoteAddress Defaultgateway
-RemoteAddress LocalSubnet,Defaultgateway
-RemoteAddress 192.168.9.120-192.168.9.130
-RemoteAddress 192.168.9.120/16
-RemoteAddress FE80::0202:B3FF:FE1E:8329
-RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340
#>
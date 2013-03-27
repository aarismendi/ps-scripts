#region ParameterSet ByName
$test = 'Remove-FirewallRule (ByName): '
Remove-FirewallRule -All

New-FirewallRule -Name foo | Out-Null
New-FirewallRule -Name foo2 | Out-Null
Remove-FirewallRule -Name foo 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Inbound | Out-Null
New-FirewallRule -Name foo -Direction Outbound | Out-Null
Remove-FirewallRule -Name foo -Direction Outbound 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Direction(Outbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain | Out-Null
New-FirewallRule -Name foo -Profile Public | Out-Null
Remove-FirewallRule -Name foo -Profile Domain 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Profile(Domain)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress Defaultgateway | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(LocalSubnet,Defaultgateway)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120-192.168.9.130)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120/16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
Remove-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Out-Null
New-FirewallRule -Name foo -Protocol udp | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 40 | Out-Null
New-FirewallRule -Name foo -Protocol 50 | Out-Null
Remove-FirewallRule -Name foo -Protocol 50 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(50)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
Remove-FirewallRule -Name foo -Protocol 'ICMPv4:3,5' 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(ICMPv4:3,5)"}
Remove-FirewallRule -All

<# netsh protocol(any) doesn't work here
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
Remove-FirewallRule -Name foo -Protocol any  -Debug
$r = @(Get-FirewallRule)
if ($r -ne $null) {$r; throw "Failed $test name(foo),Protocol(any)"}
Remove-FirewallRule -All
#>

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -LocalPort 90 | Out-Null
Remove-FirewallRule -Name foo -LocalPort 80 -Protocol tcp 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),LocalPort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80-100 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -LocalPort 90 | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp -LocalPort 80-100 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),LocalPort(80-100)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -RemotePort 100-120 | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp -RemotePort 80 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81,82 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81 | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81,82 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(80,81,82)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81,82 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -RemotePort 100-120 | Out-Null
Remove-FirewallRule -Name foo -Protocol tcp -RemotePort 100-120 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(100-120)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Program C:\notepad.exe | Out-Null
New-FirewallRule -Name foo -Program C:\ping.exe | Out-Null
Remove-FirewallRule -Name foo -Program C:\notepad.exe 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Program(C:\notepad.exe)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Service W32Time | Out-Null
New-FirewallRule -Name foo -Service Spooler | Out-Null
Remove-FirewallRule -Name foo -Service Spooler 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Service(Spooler)"}
Remove-FirewallRule -All
#endregion

#region ParameterSet ByInput
$test = 'Remove-FirewallRule (ByInput): '
Remove-FirewallRule -All

New-FirewallRule -Name foo | Out-Null
New-FirewallRule -Name bar | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(bar)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Inbound | Out-Null
New-FirewallRule -Name foo -Direction Outbound | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Direction(Outbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain | Out-Null
New-FirewallRule -Name foo -Profile Public | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Profile(Domain)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress Defaultgateway | Out-Null
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(LocalSubnet,Defaultgateway)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 | Remove-FirewallRule 
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120-192.168.9.130)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 | Remove-FirewallRule 
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(192.168.9.120/16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Remove-FirewallRule 
New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),RemoteAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol udp | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 40 | Out-Null
New-FirewallRule -Name foo -Protocol 50 | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(50)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(ICMPv4:3,5)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Remove-FirewallRule 
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(ICMPv4:10,*)"}
Remove-FirewallRule -All

<# netsh protocol(any) doesn't work here
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null
New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 10:* | Out-Null
Remove-FirewallRule -Name foo -Protocol any  -Debug
$r = @(Get-FirewallRule)
if ($r -ne $null) {$r; throw "Failed $test name(foo),Protocol(any)"}
Remove-FirewallRule -All
#>

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80 | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol tcp -LocalPort 90 | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),LocalPort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80-100 | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol tcp -LocalPort 90 | Out-Null
$r = @(Get-FirewallRule)
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),LocalPort(80-100)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80 | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol tcp -RemotePort 100-120 | Out-Null
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81,82 | Remove-FirewallRule 
New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81 | Out-Null
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(80,81,82)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80,81,82 | Out-Null
New-FirewallRule -Name foo -Protocol tcp -RemotePort 100-120 | Remove-FirewallRule 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Protocol(tcp),RemotePort(100-120)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Program C:\notepad.exe | Remove-FirewallRule 
New-FirewallRule -Name foo -Program C:\ping.exe | Out-Null
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Program(C:\notepad.exe)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Service W32Time | Out-Null
New-FirewallRule -Name foo -Service Spooler | Remove-FirewallRule 
if ($r.Count -ne 1) {$r; throw "Failed $test name(foo),Service(Spooler)"}
Remove-FirewallRule -All
#endregion
#region New-FirewallRule
$test = 'New-FirewallRule: '

New-FirewallRule -Name foo | Out-Null
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test name(foo)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo  -Description bar | Out-Null
$r = Get-FirewallRule
if ($r.Name -ne 'foo' -or $r.Description -ne 'bar') {$r; throw "Failed $test name(foo),description(bar)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Action Block | Out-Null
$r = Get-FirewallRule
if ($r.Action -ne 0) {$r; throw "Failed $test action(block)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Action Allow | Out-Null
$r = Get-FirewallRule
if ($r.Action -ne 1) {$r; throw "Failed $test action(allow)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Enabled False | Out-Null
$r = Get-FirewallRule
if ($r.Enabled -ne 0) {$r; throw "Failed $test enabled(false)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Enabled True | Out-Null
$r = Get-FirewallRule
if ($r.Enabled -ne 1) {$r; throw "Failed $test enabled(true)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Outbound | Out-Null
$r = Get-FirewallRule
if ($r.Direction -ne 2) {$r; throw "Failed $test direction(Outbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Inbound | Out-Null
$r = Get-FirewallRule
if ($r.Direction -ne 1) {$r; throw "Failed $test direction(Inbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Group test | Out-Null
$r = Get-FirewallRule
if ($r.Grouping -ne 'test') {$r; throw "Failed $test group(test)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 1) {$r; throw "Failed $test protocol(ICMPv4)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 58) {$r; throw "Failed $test protocol(ICMPv6)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 3:5 | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '3:5') {$r; throw "Failed $test Protocol(ICMPv6),icmptype(3:5)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 6) {$r; throw "Failed $test protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 17) {$r; throw "Failed $test protocol(udp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol any | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 256) {$r; throw "Failed $test protocol(any)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 16 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 16) {$r; throw "Failed $test protocol(16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType any | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '*') {$r; throw "Failed $test Protocol(ICMPv4),icmptype(any)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '3:5') {$r; throw "Failed $test Protocol(ICMPv4),icmptype(3:5)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 5:* | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '5:*') {$r; throw "Failed $test Protocol(ICMPv4),icmptype(5:*)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 5:233 | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '5:233') {$r; throw "Failed $test Protocol(ICMPv4),icmptype(5:233)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalPorts -ne '80') {$r; throw "Failed $test Protocol(tcp),LocalPort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp -LocalPort 80-4000 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalPorts -ne '80-4000') {$r; throw "Failed $test Protocol(tcp),LocalPort(80-4000)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80 | Out-Null 
$r = Get-FirewallRule
if ($r.RemotePorts -ne '80') {$r; throw "Failed $test Protocol(tcp),RemotePort(80)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp -RemotePort 40,80-1000 | Out-Null 
$r = Get-FirewallRule
if ($r.RemotePorts -ne '40,80-1000') {$r; throw "Failed $test Protocol(udp),RemotePort(40,80-1000)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceAlias 'Local Area Connection' | Out-Null 
$r = Get-FirewallRule
if ($r.Interfaces -ne 'Local Area Connection') {$r; throw "Failed $test InterfaceAlias(Local Area Connection)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceType Lan | Out-Null 
$r = Get-FirewallRule
if ($r.InterfaceTypes -ne 'Lan') {$r; throw "Failed $test InterfaceType(Lan)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceType Lan,Wireless | Out-Null 
$r = Get-FirewallRule
if ($r.InterfaceTypes -ne 'Lan,Wireless') {$r; throw "Failed $test InterfaceType(Lan,Wireless)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile All | Out-Null 
$r = Get-FirewallRule
if ($r.Profiles -ne 2147483647) {$r; throw "Failed $test Profile(all)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain,Private | Out-Null 
$r = Get-FirewallRule
if ($r.Profiles -ne 3) {$r; throw "Failed $test Profile(Domain,Private)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain | Out-Null 
$r = Get-FirewallRule
if ($r.Profiles -ne 1) {$r; throw "Failed $test Profile(Domain)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Program C:\program.exe | Out-Null 
$r = Get-FirewallRule
if ($r.ApplicationName -ne 'C:\program.exe') {$r; throw "Failed $test Program(C:\program.exe)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.120/255.255.255.255') {$r; throw "Failed $test LocalAddress(192.168.9.120)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120/16 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.0.0/255.255.0.0') {$r; throw "Failed $test LocalAddress(192.168.9.120/16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120-192.168.9.130 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.120-192.168.9.130') {$r; throw "Failed $test LocalAddress(192.168.9.120-192.168.9.130)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne 'fe80::202:b3ff:fe1e:8329-fe80::202:b3ff:fe1e:8340') {$r; throw "Failed $test LocalAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress Defaultgateway | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'Defaultgateway') {$r; throw "Failed $test RemoteAddress(Defaultgateway)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'LocalSubnet,Defaultgateway') {$r; throw "Failed $test RemoteAddress(LocalSubnet,Defaultgateway)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne '192.168.9.120-192.168.9.130') {$r; throw "Failed $test RemoteAddress(192.168.9.120-192.168.9.130)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne '192.168.0.0/255.255.0.0') {$r; throw "Failed $test RemoteAddress(192.168.9.120/16)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329 | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'fe80::202:b3ff:fe1e:8329-fe80::202:b3ff:fe1e:8329') {$r; throw "Failed $test RemoteAddress(FE80::0202:B3FF:FE1E:8329)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Out-Null 
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'fe80::202:b3ff:fe1e:8329-fe80::202:b3ff:fe1e:8340') {$r; throw "Failed $test RemoteAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Service alg | Out-Null 
$r = Get-FirewallRule
if ($r.serviceName -ne 'alg') {$r; throw "Failed $test service(alg)"}
Remove-FirewallRule -All
#endregion
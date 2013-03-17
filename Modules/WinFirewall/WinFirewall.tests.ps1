cls
$ErrorActionPreference = 'Stop'
#$_fw = New-Object -ComObject HNetCfg.FwPolicy2

Import-Module -Name "~\Documents\GitHub\ps-scripts\Modules\WinFirewall\WinFirewall.psm1" -Force -Verbose

Remove-Variable * -ErrorAction SilentlyContinue
Remove-FirewallRule -All -Verbose

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

#region Set-FirewallRule
$test = 'Set-FirewallRule: '

New-FirewallRule -Name bob | Set-FirewallRule -NewName foo
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test Set(ByInput),NewName(foo)"}
Remove-FirewallRule -All

New-FirewallRule -Name bob | Out-Null
Set-FirewallRule -Name bob -NewName foo
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test Set(ByName),NewName(foo)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo  -Description bar | Set-FirewallRule -Description bob
$r = Get-FirewallRule
if ($r.Name -ne 'foo' -or $r.Description -ne 'bob') {$r; throw "Failed $test Set(ByInput),name(foo),description(bob)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Description bar | Out-Null
Set-FirewallRule -Name foo -Description bob
$r = Get-FirewallRule
if ($r.Name -ne 'foo' -or $r.Description -ne 'bob') {$r; throw "Failed $test Set(ByName),name(foo),description(bob)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Action block | Set-FirewallRule -Action allow
$r = Get-FirewallRule
if ($r.Action -ne 1) {$r; throw "Failed $test Set(ByInput),action(allow)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Action allow | Out-Null
Set-FirewallRule -Name foo -Action block
$r = Get-FirewallRule
if ($r.Action -ne 0) {$r; throw "Failed $test Set(ByName),action(block)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Enabled False | Set-FirewallRule -Enabled True
$r = Get-FirewallRule
if ($r.Enabled -ne 1) {$r; throw "Failed $test Set(ByInput),enabled(true)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Enabled True | Out-Null
Set-FirewallRule -Name foo -Enabled False
$r = Get-FirewallRule
if ($r.Enabled -ne 0) {$r; throw "Failed $test Set(ByName),enabled(false)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Inbound | Set-FirewallRule -Direction Outbound
$r = Get-FirewallRule
if ($r.Direction -ne 2) {$r; throw "Failed $test Set(ByInput),direction(Outbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Direction Outbound | Out-Null
Set-FirewallRule -Name foo -Direction Inbound
$r = Get-FirewallRule
if ($r.Direction -ne 1) {$r; throw "Failed $test Set(ByName),direction(Inbound)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Group test | Set-FirewallRule -Group bob
$r = Get-FirewallRule
if ($r.Grouping -ne 'bob') {$r; throw "Failed $test Set(ByInput),group(test)"}
Remove-FirewallRule -All

New-FirewallRule -Name abc -Group test | Out-Null
New-FirewallRule -Name ayz -Group blah | Out-Null
Set-FirewallRule -Name a* -Group andy
$r = Get-FirewallRule
if ($r | ? {$_.Grouping -ne 'andy'}) {$r; throw "Failed $test Set(ByName),group(andy)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 | Set-FirewallRule -Protocol udp
$r = Get-FirewallRule
if ($r.Protocol -ne 17) {$r; throw "Failed $test Set(ByInput),protocol(udp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 | Out-Null
Set-FirewallRule -Name foo -Protocol 5
$r = Get-FirewallRule
if ($r.Protocol -ne 5) {$r; throw "Failed $test Set(ByName),protocol(5)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv4 | Out-Null
Set-FirewallRule -Name foo -Protocol any
$r = Get-FirewallRule
if ($r.Protocol -ne 0) {$r; throw "Failed $test Set(ByName),protocol(any)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 4 | Set-FirewallRule -Protocol tcp
$r = Get-FirewallRule
if ($r.Protocol -ne 6) {$r; throw "Failed $test Set(ByInput),protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol 4 | Out-Null
Set-FirewallRule -Name foo -Protocol tcp
$r = Get-FirewallRule
if ($r.Protocol -ne 6) {$r; throw "Failed $test Set(ByName),protocol(tcp)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 3:5 | Set-FirewallRule -IcmpType 6:*
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '6:*') {$r; throw "Failed $test Set(ByInput),Protocol(ICMPv6),icmptype(6:*)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 3:5 | Out-Null
Set-FirewallRule -Name foo -IcmpType 6:*
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '6:*') {$r; throw "Failed $test Set(ByName),Protocol(ICMPv6),icmptype(6:*)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 3:5 | Set-FirewallRule -Protocol ICMPv4 -IcmpType 100:200
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '100:200' -and $r.Protocol -ne 1) {$r; throw "Failed $test Set(ByInput),Protocol(ICMPv4),icmptype(100:200)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol ICMPv6 -IcmpType 3:5 | Out-Null
Set-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 100:200
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '100:200' -and $r.Protocol -ne 1) {$r; throw "Failed $test Set(ByName),Protocol(ICMPv4),icmptype(100:200)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Set-FirewallRule -Protocol ICMPv4 -IcmpType any
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '*' -and $r.Protocol -ne 1) {$r; throw "Failed $test Set(ByInput),Protocol(ICMPv4),icmptype(any)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp | Out-Null
Set-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType any
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '*' -and $r.Protocol -ne 1) {$r; throw "Failed $test Set(ByName),Protocol(ICMPv4),icmptype(any)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80 | Set-FirewallRule -LocalPort 100
$r = Get-FirewallRule
if ($r.LocalPorts -ne '100') {$r; throw "Failed $test Set(ByInput),Protocol(tcp),LocalPort(100)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -LocalPort 80 | Out-Null
Set-FirewallRule -Name foo -LocalPort 100
$r = Get-FirewallRule
if ($r.LocalPorts -ne '100') {$r; throw "Failed $test Set(ByName),Protocol(tcp),LocalPort(100)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp -LocalPort 80-4000 | Set-FirewallRule -LocalPort 100-1000
$r = Get-FirewallRule
if ($r.LocalPorts -ne '100-1000') {$r; throw "Failed $test Set(ByInput),Protocol(udp),LocalPort(80-4000)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp -LocalPort 80-4000 | Out-Null
Set-FirewallRule -Name foo -LocalPort 100-1000
$r = Get-FirewallRule
if ($r.LocalPorts -ne '100-1000') {$r; throw "Failed $test Set(ByName),Protocol(udp),LocalPort(100-1000)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80 | Set-FirewallRule -RemotePort 199
$r = Get-FirewallRule
if ($r.RemotePorts -ne '199') {$r; throw "Failed $test Set(ByInput),Protocol(tcp),RemotePort(199)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol tcp -RemotePort 80 | Out-Null
Set-FirewallRule -Name foo -RemotePort 199
$r = Get-FirewallRule
if ($r.RemotePorts -ne '199') {$r; throw "Failed $test Set(ByName),Protocol(tcp),RemotePort(199)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Protocol udp -RemotePort 40,80-1000 | Set-FirewallRule -RemotePort 50,1000-2000
$r = Get-FirewallRule
if ($r.RemotePorts -ne '50,1000-2000') {$r; throw "Failed $test Protocol(udp),RemotePort(50,1000-2000)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceAlias 'Local Area Connection' | Set-FirewallRule -InterfaceAlias 'Local Area Connection 2'
$r = Get-FirewallRule
if ($r.Interfaces -ne 'Local Area Connection 2') {$r; throw "Failed $test InterfaceAlias(Local Area Connection 2)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceType Lan | Set-FirewallRule -InterfaceType Wireless
$r = Get-FirewallRule
if ($r.InterfaceTypes -ne 'Wireless') {$r; throw "Failed $test InterfaceType(Wireless)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -InterfaceType Lan,Wireless | Set-FirewallRule -InterfaceType RemoteAccess
$r = Get-FirewallRule
if ($r.InterfaceTypes -ne 'RemoteAccess') {$r; throw "Failed $test InterfaceType(RemoteAccess)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile All | Set-FirewallRule -Profile Domain
$r = Get-FirewallRule
if ($r.Profiles -ne 1) {$r; throw "Failed $test Profile(Domain)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain,Private | Set-FirewallRule -Profile All
$r = Get-FirewallRule
if ($r.Profiles -ne 2147483647) {$r; throw "Failed $test Profile(All)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Profile Domain | Set-FirewallRule -Profile Private
$r = Get-FirewallRule
if ($r.Profiles -ne 2) {$r; throw "Failed $test Profile(Private)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Program C:\program.exe | Set-FirewallRule -Program C:\new_program.exe
$r = Get-FirewallRule
if ($r.ApplicationName -ne 'C:\new_program.exe') {$r; throw "Failed $test Program(C:\new_program.exe)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120 | Set-FirewallRule -LocalAddress 192.168.9.121
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.121/255.255.255.255') {$r; throw "Failed $test LocalAddress(192.168.9.121)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120/16 | Set-FirewallRule -LocalAddress 192.168.9.120/8
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.0.0.0/255.0.0.0') {$r; throw "Failed $test LocalAddress(192.168.9.120/8)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress 192.168.9.120-192.168.9.130 | Set-FirewallRule -LocalAddress 192.168.9.140-192.168.9.150
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.140-192.168.9.150') {$r; throw "Failed $test LocalAddress(192.168.9.140-192.168.9.150)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -LocalAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | 
    Set-FirewallRule -LocalAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8341
$r = Get-FirewallRule
if ($r.LocalAddresses -ne 'fe80::202:b3ff:fe1e:8329-fe80::202:b3ff:fe1e:8341') {$r; throw "Failed $test LocalAddress(FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8341)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress Defaultgateway | Set-FirewallRule -RemoteAddress dns
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'dns') {$r; throw "Failed $test RemoteAddress(dns)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress LocalSubnet,Defaultgateway | Set-FirewallRule -RemoteAddress WINS,DNS
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'DNS,WINS') {$r; throw "Failed $test RemoteAddress(WINS,DNS)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120-192.168.9.130 | Set-FirewallRule -RemoteAddress 192.168.9.140-192.168.9.150
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne '192.168.9.140-192.168.9.150') {$r; throw "Failed $test RemoteAddress(192.168.9.140-192.168.9.150)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress 192.168.9.120/16 | Set-FirewallRule -RemoteAddress 192.168.9.120/8
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne '192.0.0.0/255.0.0.0') {$r; throw "Failed $test RemoteAddress(192.168.9.120/8)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329 | Set-FirewallRule -RemoteAddress FE80::0202:B3FF:FE1E:8340
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'fe80::202:b3ff:fe1e:8340-fe80::202:b3ff:fe1e:8340') {$r; throw "Failed $test RemoteAddress(FE80::0202:B3FF:FE1E:8340)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -RemoteAddress FE80::0202:B3FF:FE1E:8329-FE80::0202:B3FF:FE1E:8340 | Set-FirewallRule -RemoteAddress FE80::0202:B3FF:FE1E:8350-FE80::0202:B3FF:FE1E:8360
$r = Get-FirewallRule
if ($r.RemoteAddresses -ne 'fe80::202:b3ff:fe1e:8350-fe80::202:b3ff:fe1e:8360') {$r; throw "Failed $test RemoteAddress(FE80::0202:B3FF:FE1E:8350-FE80::0202:B3FF:FE1E:8360)"}
Remove-FirewallRule -All

New-FirewallRule -Name foo -Service alg | Set-FirewallRule -Service vds
$r = Get-FirewallRule
if ($r.serviceName -ne 'vds') {$r; throw "Failed $test service(vds)"}
Remove-FirewallRule -All
#endregion

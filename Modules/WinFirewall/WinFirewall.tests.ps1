cls
$ErrorActionPreference = 'Stop'
#$_fw = New-Object -ComObject HNetCfg.FwPolicy2

Import-Module -Name "~\Documents\GitHub\ps-scripts\Modules\WinFirewall\WinFirewall.psm1" -Force -Verbose

Remove-FirewallRule -All -Verbose

#region New-FirewallRule
$test = 'new rule with name'
New-FirewallRule -Name foo | Out-Null
$r = Get-FirewallRule
if ($r.Name -ne 'foo') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'new rule with name,description'
New-FirewallRule -Name foo  -Description bar | Out-Null
$r = Get-FirewallRule
if ($r.Name -ne 'foo' -or $r.Description -ne 'bar') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'action(block)'
New-FirewallRule -Name foo -Action Block | Out-Null
$r = Get-FirewallRule
if ($r.Action -ne 0) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'action(allow)'
New-FirewallRule -Name foo -Action Allow | Out-Null
$r = Get-FirewallRule
if ($r.Action -ne 1) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'enabled(false)'
New-FirewallRule -Name foo -Enabled False | Out-Null
$r = Get-FirewallRule
if ($r.Enabled -ne 0) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'enabled(true)'
New-FirewallRule -Name foo -Enabled True | Out-Null
$r = Get-FirewallRule
if ($r.Enabled -ne 1) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'direction(Outbound)'
New-FirewallRule -Name foo -Direction Outbound | Out-Null
$r = Get-FirewallRule
if ($r.Direction -ne 2) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'direction(Inbound)'
New-FirewallRule -Name foo -Direction Inbound | Out-Null
$r = Get-FirewallRule
if ($r.Direction -ne 1) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'group(test)'
New-FirewallRule -Name foo -Group test | Out-Null
$r = Get-FirewallRule
if ($r.Grouping -ne 'test') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(ICMPv4)'
New-FirewallRule -Name foo -Protocol ICMPv4 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 1) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(ICMPv6)'
New-FirewallRule -Name foo -Protocol ICMPv6 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 58) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(tcp)'
New-FirewallRule -Name foo -Protocol tcp | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 6) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(udp)'
New-FirewallRule -Name foo -Protocol udp | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 17) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(any)'
New-FirewallRule -Name foo -Protocol any | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 256) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'protocol(16)'
New-FirewallRule -Name foo -Protocol 16 | Out-Null 
$r = Get-FirewallRule
if ($r.Protocol -ne 16) {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'icmptype(any)'
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType any | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '*') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'icmptype(3:5)'
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 3:5 | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '3:5') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'icmptype(5:*)'
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 5:* | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '5:*') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'icmptype(5:233)'
New-FirewallRule -Name foo -Protocol ICMPv4 -IcmpType 5:233 | Out-Null 
$r = Get-FirewallRule
if ($r.IcmpTypesAndCodes -ne '5:233') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'icmptype(Local Area Connection)'
New-FirewallRule -Name foo -InterfaceAlias 'Local Area Connection' | Out-Null 
$r = Get-FirewallRule
if ($r.Interfaces -ne 'Local Area Connection') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'InterfaceType(Lan)'
New-FirewallRule -Name foo -InterfaceType Lan | Out-Null 
$r = Get-FirewallRule
if ($r.InterfaceTypes -ne 'Lan') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'LocalAddress(192.168.9.120)'
New-FirewallRule -Name foo -LocalAddress 192.168.9.120 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.120/255.255.255.255') {$r; throw "Failed $test"}
Remove-FirewallRule -All

$test = 'LocalAddress(192.168.9.120/16)'
New-FirewallRule -Name foo -LocalAddress 192.168.9.120/8 | Out-Null 
$r = Get-FirewallRule
if ($r.LocalAddresses -ne '192.168.9.120/255.255.0.0') {$r; throw "Failed $test"}
Remove-FirewallRule -All
#endregion

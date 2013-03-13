#region Internal
function Get-FirewallConfigObject {
	# http://msdn.microsoft.com/en-us/library/windows/desktop/aa365309(v=vs.85).aspx
	return New-Object -ComObject HNetCfg.FwPolicy2
}

function Get-FirewallProfileBitmask {
	[cmdletbinding()] param (
		[parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public', 'current', 'all')] [string[]] $Name)
	$fw = Get-FirewallConfigObject
	$profile_enum = @{domain = 1; private = 2; public = 4; all = 2147483647; current = $fw.CurrentProfileTypes; selected = 0}
	$Name | % {$profile_enum.selected = $profile_enum.selected -bor $profile_enum[$_]}
	return $profile_enum.selected
}

function Test-IsAdmin {  
	return (([Security.Principal.WindowsPrincipal] `
				[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
					[Security.Principal.WindowsBuiltInRole] "Administrator"))	
}

function Test-IsSupportedOS {
	return ([Environment]::OSVersion.Version -ge (new-object 'Version' 6,0))
}

function Initialize {
	if (-not (Test-IsAdmin)) {
		throw 'This requires running as admin.'
	}
	if (-not (Test-IsSupportedOS)) {
		throw 'This requires Windows Vista/2008 or greater.'
	}
}

function Set-Netsh {
	[cmdletbinding()] param (
		[parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Name,
		[string] $Operation,
		[string] $Parameter,
		[string] $Value, 
		[string[]] $Arguments
	)
	if ($Arguments) {
		$arguments = ('advfirewall', 'set') + $Arguments
		Write-Debug ("Running netsh.exe with arguments: " + $arguments)
		$output = [string[]] (& netsh.exe $arguments 2>&1)
		if ($LASTEXITCODE -ne 0) {throw $output[1]} # TODO do this in a cmdlety way

	} else {
		$netsh_profile_map = @{domain = 'domainprofile'; private = 'privateprofile'; public = 'publicprofile'}
		foreach ($profile_name in $Name) {
			$netsh_profile_name = $netsh_profile_map[$profile_name]
			$arguments = 'advfirewall', 'set', $netsh_profile_name, $Operation, $Parameter, $Value
			Write-Debug ("Running netsh.exe with arguments: " + $arguments)
			$output = [string[]] (& netsh.exe $arguments 2>&1)
			if ($LASTEXITCODE -ne 0) {throw $output[1]} # TODO do this in a cmdlety way
		}
	}
}
#endregion

#region Exports
#region Profile ############################################################################
function Get-FirewallProfile {
	# TODO create
}

function Set-FirewallProfile {
	[cmdletbinding(DefaultParameterSetName="ByName")] param (
		[parameter(ParameterSetName="ByName", Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Name,
		[parameter()] [ValidateSet('True', 'False')] [string] $Enabled,
		[parameter()] [ValidateSet('allow', 'block', 'block_all')] [string] $DefaultInboundAction,
		[parameter()] [ValidateSet('allow', 'block')] [string] $DefaultOutboundAction,
		[parameter()] [string] $LogFileName,
		[parameter()] [string] $LogMaxSizeKilobytes,
		[parameter()] [ValidateSet('True', 'False')] [string] $LogDroppedPackets,
		[parameter()] [ValidateSet('True', 'False')] [string] $LogConnections,
		[parameter()] [ValidateSet('True', 'False')] [string] $DisableNotifications,
		[parameter()] [ValidateSet('True', 'False')] [string] $DisableUnicastResponse
	)
	process {
		Initialize
		$fw = Get-FirewallConfigObject
		$action_enum = @{block = 0; allow = 1}
		$netsh_action_map = @{True = 'enable'; False = 'disable'}

		foreach ($profile_name in $Name) {
			$profile_mask = Get-FirewallProfileBitmask -Name $profile_name

			if ($PSBoundParameters.ContainsKey('Enabled')) {$fw.FirewallEnabled($profile_mask) = [boolean]::Parse($Enabled)}
			
			if ($PSBoundParameters.ContainsKey('DefaultInboundAction')) {
				if ($DefaultInboundAction -eq 'block_all') {$fw.BlockAllInboundTraffic($profile_mask) = $true
				} else {
					$fw.BlockAllInboundTraffic($profile_mask) = $false
					$fw.DefaultInboundAction($profile_mask) = $action_enum[$DefaultInboundAction]}
			}
			
			if ($PSBoundParameters.ContainsKey('DefaultOutboundAction')) {
				$fw.DefaultOutboundAction($profile_mask) = $action_enum[$DefaultOutboundAction]}
			
			if ($PSBoundParameters.ContainsKey('LogFileName')) {
				if ($LogFileName -eq 'default') {$LogFileName = '%systemroot%\system32\LogFiles\Firewall\pfirewall.log'}
				Set-Netsh -Name $profile_name -Operation 'logging' -Parameter filename -Value $LogFileName}
			
			if ($PSBoundParameters.ContainsKey('LogMaxSizeKilobytes')) {
				if ($LogMaxSizeKilobytes -eq 'default') {$LogMaxSizeKilobytes = 4096}
				Set-Netsh -Name $profile_name -Operation 'logging' -Parameter maxfilesize -Value $LogMaxSizeKilobytes}
			
			if ($PSBoundParameters.ContainsKey('LogDroppedPackets')) {
				Set-Netsh -Name $profile_name -Operation 'logging' -Parameter droppedconnections -Value ($netsh_action_map[$LogDroppedPackets])}
			
			if ($PSBoundParameters.ContainsKey('LogConnections')) {
				Set-Netsh -Name $profile_name -Operation 'logging' -Parameter allowedconnections -Value ($netsh_action_map[$LogConnections])}
			
			if ($PSBoundParameters.ContainsKey('DisableNotifications')) {
				$fw.NotificationsDisabled($profile_mask) = [boolean]::Parse($DisableNotifications)}
			
			if ($PSBoundParameters.ContainsKey('DisableUnicastResponse')) {
				$fw.UnicastResponsesToMulticastBroadcastDisabled($profile_mask) = [boolean]::Parse($DisableUnicastResponse)}
		}
	}
}
#endregion Profile #########################################################################

#region Rules ##############################################################################
function Get-FirewallRule {
	# TODO support some filters
	(Get-FirewallConfigObject).Rules
}

function New-FirewallRule {
	[cmdletbinding()] param ( 
		[parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateScript({$_ -notmatch '^all$|\|'})] [string] $Name, 
		[parameter(ParameterSetName="Make")] [ValidateScript({-not $_.Contains('|')})] [string] $Description,
		[parameter(ParameterSetName="Make", Mandatory=$true)] [ValidatePattern('^\d{1,5}$|^\d{1,5}-\d{1,5}$')] [string[]] $Port, # Supported Format: '80','443','8000-8009'
		[parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('tcp','udp')] [string] $Protocol,
		[parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('in','out')] [string] $Direction,
		[parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('allow','block')] [string] $Action,
		[parameter(ParameterSetName="Make")] [ValidateSet('domain','private','public','current','all')] [string[]] $Profile = 'all',
		[parameter(ParameterSetName="Make")] [ValidateSet('RemoteAccess','Wireless','Lan','All')] [string[]] $InterfaceTypes = 'All',
		[parameter(ParameterSetName="Make")] [switch] $Disabled,
		[parameter(ParameterSetName="Make")] [string] $GroupingName,
		[parameter(ParameterSetName="Make")] [string] $ApplicationPath,
		[parameter(ParameterSetName="Make")] [ValidateScript({ ((gsv | Select -Exp Name) + '*') -contains $_})] [string] $ServiceShortName,
		[parameter(ParameterSetName="Make")] [string[]] $InterfaceNames,
		[parameter(ParameterSetName="Made")] $FirewallRuleObject
	)
	Initialize
	$fw = Get-FirewallConfigObject
	switch ($PSCmdlet.ParameterSetName) {
		'Make' {
				$rule = New-Object -ComObject HNetCfg.FWRule 

				$action_enum = @{allow = 1; block = 0}
				$direction_enum = @{in = 1; out = 2}
				$protocol_enum = @{ICMPv4 = 1; ICMPv6 = 58; tcp = 6; udp = 17}
				$profile_mask = Get-FirewallProfileBitmask -Name $Profile

				# reference: http://msdn.microsoft.com/en-us/library/windows/desktop/aa365344(v=vs.85).aspx
				# reference: http://msdn.microsoft.com/en-us/library/bb945065.aspx
				# GUI: General
				$rule.Name = $Name ; $rule.Description = $Description
				$rule.Enabled = (-not $Disabled)
				$rule.Action = $action_enum[$Action]
				# GUI: Programs and Services
				if ($ApplicationPath) { $rule.ApplicationName = $ApplicationPath }
				if ($ServiceShortName) { $rule.serviceName = $ServiceShortName }
				# GUI: Protocols and Ports
				$rule.Protocol = $protocol_enum[$Protocol]
				$rule.LocalPorts = ($Port -join ',') # protocal must be set first!
				$rule.RemotePorts = '*'
				# GUI: Scope
				$rule.LocalAddresses = '*'
				$rule.RemoteAddresses = '*'
				# GUI: Advanced
				$rule.Profiles = $profile_mask
				$rule.InterfaceTypes = ($InterfaceTypes -join ',')
				# not used $rule.EdgeTraversal = $false ; $rule.EdgeTraversalOptions = $null
				# GUI: N/A
				$rule.Direction = $direction_enum[$Direction]
				$rule.Grouping = $GroupingName
				$nic_names = @(); $nic_names += $InterfaceNames
				if ($InterfaceNames) { $rule.Interfaces = $nic_names }
				# not used: $rule.IcmpTypesAndCodes = $null
		
		}
		'Made' {$rule = $FirewallRuleObject}
	}

	$rule_print = ($rule | gm -MemberType Property | % {"{0} : {1}" -f $_.Name, $rule."$($_.Name)"}) -join "`n"
	Write-Verbose ("Adding firewall rule:`n" + $rule_print)
	
	try {
		$fw.Rules.Add($rule)
		return $rule
	} catch {
		$PSCmdlet.WriteError($_)
	}
}

function Remove-FirewallRule {
	[cmdletbinding(DefaultParameterSetName="ByName", SupportsShouldProcess=$True, ConfirmImpact='Medium')] param ( 
		[parameter(Mandatory=$true, ParameterSetName="ByName", ValueFromPipeline=$true)] [string[]] $Name,
		[parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true)] $InputObject
	)
	begin {Initialize; $fw = Get-FirewallConfigObject}
	process {
		if ($PSCmdlet.ParameterSetName -eq 'ByName') {
			$pattern = New-Object System.Management.Automation.WildcardPattern $Name, 'Compiled,IgnoreCase'
			$rules = $fw.Rules | ? {$pattern.IsMatch($_.Name)}
			$rules | % {
				Write-Verbose ("Removing firewall rule: " + $_.Name)
				if ($PSCmdlet.ShouldProcess($_.Name)) {$fw.Rules.Remove($_.Name)}
			 }
		} else {
			$rule = $InputObject
			$arguments = @()
			$direction_enum = @{1 = 'in'; 2 = 'out'}
			$rule | gm -MemberType Property | Select -ExpandProperty Name | ? {$rule."$_"} | % {
				switch ($_) {
					Direction {$arguments += ('dir={0}' -f $direction_enum[$rule."$_"])}
					Name {$arguments += ('name={0}' -f $rule.Name)}
					Profiles {$profile_enum = @{1='domain';2='private';4='public'}
						$profile_mask = $rule."$_"
						$profile_list = ($profile_enum.Keys | ? {$profile_mask -band $_} | % {$profile_enum[$_]}) -join ','
						$arguments += ('profile={0}' -f $profile_list)} # [profile=public|private|domain|any[,...]]
					RemoteAddresses {$arguments += ('remoteip={0}' -f $rule."$_".Replace('*', 'any'))} # [remoteip=any|localsubnet|dns|dhcp|wins|defaultgateway|<IPv4 address>|<IPv6 address>|<subnet>|<range>|<list>]
					RemotePorts {$arguments += ('remoteport={0}' -f $rule."$_".Replace('*', 'any'))} # [remoteport=0-65535|<port range>[,...]|any]
					ApplicationName {$arguments += ('program={0}' -f $rule."$_")} # [program=<program path>]
					serviceName {$arguments += ('service={0}' -f $rule."$_")}  # [service=<service short name>|any]
					Protocol {$arguments += ('protocol={0}' -f $rule."$_")} # [protocol=0-255|icmpv4|icmpv6|icmpv4:type,code|icmpv6:type,code|tcp|udp|any]
					LocalPorts {$arguments += ('localport={0}' -f $rule."$_".Replace('*', 'any'))} # [localport=0-65535|<port range>[,...]|RPC|RPC-EPMap|any]
				}
			}
			$msg = ('Removing firewall rule with: netsh.exe ' + ($arguments -join ' '))
			if ($PSCmdlet.ShouldProcess(($arguments -join ' '))) {
				Write-Debug $msg
				$arguments = ('advfirewall', 'firewall', 'delete', 'rule') + $arguments
				$output = [string] (& netsh.exe $arguments 2>$1)
				if ($LASTEXITCODE -ne 0) {
					throw $output
				}
			}
		}
	}
}

function Set-FirewallRule {
	[cmdletbinding(DefaultParameterSetName="ByName")] param ( 
		[parameter(HelpMessage='The action to take for traffic which matches this rule.')]
		[ValidateSet('allow', 'block')] [string] $Action = 'allow',
		
		[parameter(HelpMessage='Information about the firewall rule.')] 
		[ValidateScript({-not $_.Contains('|')})] [string] $Description,
		
		[parameter(HelpMessage='Which direction of traffic to match with this rule.')]
		[ValidateSet('in','out')] [string] $Direction,
		
		[parameter(HelpMessage='Whether the rule is in effect or not.')] 
		[ValidateSet('True', 'False')] [string] $Enabled,
		
		[parameter(HelpMessage='Display name (alias) of the network interface that applies to the traffic.')] 
		[string[]] $InterfaceAlias,
		
		[parameter(HelpMessage='Only the specified network connection types are subject to this rule.')] 
		[ValidateSet('RemoteAccess','Wireless','Lan','All')] [string[]] $InterfaceType,
		
		[parameter(Mandatory=$true, ParameterSetName="ByName", 
			HelpMessage='Only the matching firewall rules of the indicated name are modified.')] 
		[ValidateScript({$_ -notmatch '^all$|\|'})] [string[]] $Name,

		[parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true, 
			HelpMessage='Firewall wall rule object to modify.')]
		$InputObject,
		
		[parameter(HelpMessage='Change the name of the firewall rule.')]
		[ValidateScript({$_ -notmatch '^all$|\|'})] [string] $NewName,
		
		[parameter(HelpMessage='Specifies one or more profiles to which the rule is assigned.')]
		[ValidateSet('domain','private','public','current','all')] [string[]] $Profile,
		
		[parameter(HelpMessage='Specifies the path and file name of the program for which the rule allows traffic.')] 
		[string] $Program,
		
		[parameter(HelpMessage='Specifies that network packets with matching protocol match this rule.')]
		[ValidateSet('tcp','udp', 'ICMPv4', 'ICMPv6')] [string] $Protocol,
		
		[parameter(HelpMessage='Specifies that network packets with matching IP addresses match this rule.')]
		[ValidateScript({($_ -match '\*|Defaultgateway|DHCP|DNS|WINS|LocalSubnet') -or ([System.Net.IPAddress]::Parse($_))})]
		[string[]] $RemoteAddress,
		
		[parameter(HelpMessage='Specifies that network packets with matching IP port numbers match this rule.')]
		[ValidatePattern('^\d{1,5}$|^\d{1,5}-\d{1,5}$')] [string[]] $RemotePort,
		
		[parameter(HelpMessage='Specifies the short name of a Windows service to which the firewall rule applies.')]
		[ValidateScript({((gsv|Select -Exp Name) + '*') -contains $_})] [string] $Service
	)

	begin {Initialize; $fw = Get-FirewallConfigObject}
	process {
		if ($PSCmdlet.ParameterSetName -eq 'ByName') {
			$pattern = New-Object System.Management.Automation.WildcardPattern $Name, 'Compiled,IgnoreCase'
			$rules = $fw.Rules | ? {$pattern.IsMatch($_.Name)}
		} else {
			$rules = $InputObject
		}

		$action_enum = @{allow = 1; block = 0}
		$direction_enum = @{in = 1; out = 2}
		$protocol_enum = @{ICMPv4 = 1; ICMPv6 = 58; tcp = 6; udp = 17}
		$v_msg = "Setting {0} '{1}' on rule named: {2}."

		if ($PSBoundParameters.ContainsKey('NewName')) {
			$rules | % {Write-Verbose ($v_msg -f 'name', $NewName, $_.Name); $_.Name = $NewName}}
		if ($PSBoundParameters.ContainsKey('Description')) {
			$rules | % {Write-Verbose ($v_msg -f 'description', $Description, $_.Name); $_.Description = $Description}}
		if ($PSBoundParameters.ContainsKey('Action')) {
			$rules | % {Write-Verbose ($v_msg -f 'action', $Action, $_.Name); $_.Action = $action_enum[$Action]}}
		if ($PSBoundParameters.ContainsKey('Direction')) {
			$rules | % {Write-Verbose ($v_msg -f 'direction', $Direction, $_.Name); $_.Direction = $direction_enum[$Direction]}}
		if ($PSBoundParameters.ContainsKey('Enabled')) {
			$rules | % {Write-Verbose ($v_msg -f 'enabled', $Enabled, $_.Name); $_.Enabled = ([bool]::Parse($Enabled))}}
		if ($PSBoundParameters.ContainsKey('InterfaceAlias')) {
			$nic_names = @(); $nic_names += $InterfaceAlias # Must be an array.
			$rules | % {Write-Verbose ($v_msg -f 'interface alias', ($nic_names -join ', '), $_.Name); $_.Interfaces =$nic_names}}
		if ($PSBoundParameters.ContainsKey('InterfaceType')) {
			$interface_list = $InterfaceType -join ','
			$rules | % {Write-Verbose ($v_msg -f 'interface type', $interface_list, $_.Name); $_.InterfaceTypes = $interface_list}}
		if ($PSBoundParameters.ContainsKey('Profile')) {
			$profile_mask = Get-FirewallProfileBitmask -Name $Profile
			$rules | % {Write-Verbose ($v_msg -f 'profile', ($Profile -join ', '), $_.Name); $_.Profiles = $profile_mask}}
		if ($PSBoundParameters.ContainsKey('Program')) {
			$rules | % {Write-Verbose ($v_msg -f 'program', $Program, $_.Name); $_.ApplicationName = $Program}}
		if ($PSBoundParameters.ContainsKey('Protocol')) {
			$rules | % {Write-Verbose ($v_msg -f 'protocol', $Protocol, $_.Name); $_.Protocol = $protocol_enum[$Protocol]}}
		if ($PSBoundParameters.ContainsKey('RemoteAddress')) {
			$rules | % {Write-Verbose ($v_msg -f 'remote address', ($RemoteAddress -join ', '), $_.Name);
						$_.RemoteAddresses = ($RemoteAddress -join ',')}}
		if ($PSBoundParameters.ContainsKey('RemotePort')) {
			$rules | % {Write-Verbose ($v_msg -f 'remote port', ($RemotePort -join ', '), $_.Name); $_.RemotePorts = ($RemotePort -join ',')}}
		if ($PSBoundParameters.ContainsKey('Service')) {
			$rules | % {Write-Verbose ($v_msg -f 'service', $Service, $_.Name); $_.serviceName = $Service}}
	}
}
#endregion Rules ###########################################################################
#endregion Exports

Export-ModuleMember -Function Get-FirewallRule, 
		Get-FirewallProfile, Set-FirewallProfile,
		New-FirewallRule, Remove-FirewallRule, Set-FirewallRule, Disable-FirewallRule, Enable-FirewallRule

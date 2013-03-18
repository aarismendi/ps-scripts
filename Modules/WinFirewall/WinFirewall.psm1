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
		[string] $Value
	)

	$netsh_profile_map = @{domain = 'domainprofile'; private = 'privateprofile'; public = 'publicprofile'}
	foreach ($profile_name in $Name) {
		$netsh_profile_name = $netsh_profile_map[$profile_name]
		$arguments = 'advfirewall', 'set', $netsh_profile_name, $Operation, $Parameter, $Value
		Write-Debug ("Running netsh.exe with arguments: " + $arguments)
		$output = [string[]] (& netsh.exe $arguments 2>&1)
		if ($LASTEXITCODE -ne 0) {throw $output[1]} # TODO do this in a cmdlety way
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
	[cmdletbinding(SupportsShouldProcess=$True)] 
    param (
        [ValidateSet('Allow','Block')] [string] 
            $Action,
        [ValidateScript( {-not $_.Contains('|')} )] [string] 
            $Description,
        [ValidateSet('Inbound','Outbound')] [string] 
            $Direction,
        [ValidateSet('True','False')] [string] 
            $Enabled = 'True',
        [string] 
            $Group,
        [ValidateScript( # Can be [0-255|*]:[0-255|*] | any
            {$_ -eq 'any' -or ((0..255 -contains $_.Split(':')[0]) -and (0..255+'*' -contains $_.Split(':')[1]))} 
        )] [string[]] 
            $IcmpType,
        [string[]] 
            $InterfaceAlias,
        [ValidateSet('RemoteAccess','Wireless','Lan','All')] [string[]] 
            $InterfaceType,
        [ValidateScript( # Can be IPv[4|6][/cidr]|range
            {$_.Split('-') | % { [System.Net.IPAddress] ($_ -replace '/\d+$','') }}
        )] [string[]] 
            $LocalAddress,
        [ValidateScript( # Protocol must be TCP|UDP. Can be 1-65535-1-65535 | 1-65535
            {($_.Split('-') | ? {1..65535 -contains $_}).Count -eq $_.Split('-').Count}
        )] [string[]] 
            $LocalPort,
	    [parameter(Mandatory=$true)] [ValidateScript( {'|','all' -notcontains $_} )] [string] 
            $Name,
        [ValidateSet('Domain','Private','Public','Current','All')] [string[]] 
            $Profile,
        [string] 
            $Program,
        [ValidateScript( 
            {(0..255 -contains $_) -or ('Any','TCP','UDP','ICMPv4','ICMPv6' -contains $_)}
        )] [string] 
            $Protocol,
        [ValidateScript( # Can be any|Defaultgateway|DHCP|WINS|LocalSubnet|IPv(4|6)addr or range|netmask or IPv4addr/cidr
            {((New-Object -ComObject HNetCfg.FWRule).RemoteAddresses = ($_ -replace 'any', '*')) -ne $null}
        )] [string[]] 
            $RemoteAddress,
        [ValidateScript( # Same as LocalPort
            {($_.Split('-') | ? {1..65535 -contains $_}).Count -eq $_.Split('-').Count}
        )] [string[]] 
            $RemotePort,
	    [ValidateScript(
            {((gsv | Select -Exp Name) + '*') -contains $_}
        )] [string] 
            $Service	    
	)
    begin {Initialize; $fw = Get-FirewallConfigObject}
    process {
  	    $action_enum = @{Allow = 1; Block = 0}
	    $direction_enum = @{Inbound = 1; Outbound = 2}
	    $protocol_enum = @{any = 256; ICMPv4 = 1; ICMPv6 = 58; TCP = 6; UDP = 17}
	
        $rule = New-Object -ComObject HNetCfg.FWRule 
	    # doc: http://msdn.microsoft.com/en-us/library/windows/desktop/aa365344(v=vs.85).aspx
	    # doc: http://msdn.microsoft.com/en-us/library/bb945065.aspx
	    $rule.Name = $Name
        $rule.Enabled = ([bool]::Parse($Enabled))
        if ($PSBoundParameters.ContainsKey('Description'))      {$rule.Description = $Description}
	    if ($PSBoundParameters.ContainsKey('Action'))           {$rule.Action = $action_enum[$Action]}
	    if ($PSBoundParameters.ContainsKey('Program'))          {$rule.ApplicationName = $Program}
	    if ($PSBoundParameters.ContainsKey('Service'))          {$rule.serviceName = $Service}
	    if ($PSBoundParameters.ContainsKey('Protocol'))         {
            $i = 0
            if ([int]::TryParse($Protocol, [ref] $i)) {$rule.Protocol = $i}
            else {$rule.Protocol = $protocol_enum[$Protocol]}
        }
	    if ($PSBoundParameters.ContainsKey('LocalPort'))        {
            if ($Protocol -match '^TCP$|^UDP$') {$rule.LocalPorts = ($LocalPort -join ',')}
            else {Write-Error "The Protocol parameter must be specified as TCP or UDP when specifiying the LocalPort parameter"}
        } #LocalPort - requires Protocol to be TCP|UDP
        if ($PSBoundParameters.ContainsKey('IcmpType'))         {$rule.IcmpTypesAndCodes = $IcmpType.Replace('any','*')}
	    if ($PSBoundParameters.ContainsKey('RemotePort'))       {
            if ($Protocol -match '^TCP$|^UDP$') {$rule.RemotePorts = ($RemotePort -join ',')}
            else {Write-Error "The Protocol parameter must be specified as TCP or UDP when specifiying the RemotePort parameter"}
        }
	    if ($PSBoundParameters.ContainsKey('LocalAddress'))     {$rule.LocalAddresses = ($LocalAddress -join ',')}
	    if ($PSBoundParameters.ContainsKey('RemoteAddress'))    {$rule.RemoteAddresses = ($RemoteAddress -join ',') -replace 'any', '*'}
	    if ($PSBoundParameters.ContainsKey('InterfaceType'))    {$rule.InterfaceTypes = ($InterfaceType -join ',')}
	    if ($PSBoundParameters.ContainsKey('Direction'))        {$rule.Direction = $direction_enum[$Direction]}
	    if ($PSBoundParameters.ContainsKey('Group'))            {$rule.Grouping = $Group}
    
        if ($PSBoundParameters.ContainsKey('InterfaceAlias')) {	
            $nic_names = @(); $nic_names += $InterfaceAlias
	        $rule.Interfaces = $nic_names 
        }
        if ($PSBoundParameters.ContainsKey('Profile')) {
            $profile_mask = Get-FirewallProfileBitmask -Name $Profile
            $rule.Profiles = $profile_mask
        }

	    $rule_print = ($rule | gm -MemberType Property | Select -ExpandProperty Name | 
                            ? {$rule."$_"} | % {"{0}={1}" -f $_, $rule."$_"}) -join " "
        Write-Debug $rule_print
        if ($PSCmdlet.ShouldProcess($rule_print)) {
	        try {
                $fw.Rules.Add($rule)
                $PSCmdlet.WriteObject($rule)
            } catch {$PSCmdlet.WriteError($_)}
        }
    }
}
function Remove-FirewallRule {
	[cmdletbinding(DefaultParameterSetName="ByAttribute", SupportsShouldProcess=$True, ConfirmImpact='Medium')] 
    param ( 
        [parameter(ParameterSetName="ByAttribute")] [string[]] 
            $Name,
        [parameter(ParameterSetName="ByAttribute")] [switch] 
            $All,
        [parameter(ParameterSetName="ByAttribute")] [ValidateSet('Inbound','Outbound')] [string] 
            $Direction,
        [parameter(ParameterSetName="ByAttribute")] [ValidateSet('Domain','Private','Public')] [string[]] 
            $Profile,
        [parameter(ParameterSetName="ByAttribute")] 
        [ValidateScript({((New-Object -ComObject HNetCfg.FWRule).RemoteAddresses = ($_ -join ',')) -ne $null})] [string[]] 
            $RemoteAddress,
        [parameter(ParameterSetName="ByAttribute")]  [ValidateScript( 
            {(0..255 -contains $_) -or ($_ -match '^icmpv4$|^icmpv6$|^icmpv4:\d{1,3},\d{1,3}$|^icmpv6:\d{1,3},\d{1,3}$|^tcp$|^udp$|^any$')}
        )] [string] 
            $Protocol,
        [parameter(ParameterSetName="ByAttribute")] [ValidatePattern('^\d{1,5}$|^\d{1,5}-\d{1,5}$')] [string[]] 
            $LocalPort,
        [parameter(ParameterSetName="ByAttribute")] [ValidatePattern('^\d{1,5}$|^\d{1,5}-\d{1,5}$')] [string[]] 
            $RemotePort,
        [parameter(ParameterSetName="ByAttribute")] [string] 
            $Program,
        [parameter(ParameterSetName="ByAttribute")] [ValidateScript({((gsv|Select -Exp Name) + '*') -contains $_})] [string] 
            $Service,
        [parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true)] 
            $InputObject
    )
	begin {Initialize; $fw = Get-FirewallConfigObject}
    process {
        <#  
            netsh supports the following attributes for selecting rules to remove. 
            netsh is more powerful than the COM API for rule removal.
            name=all|<string>
            [dir=in|out]
            [profile=public|private|domain|any[,...]]
            [program=<program path>]
            [service=<service short name>|any]
            [localip=any|<IPv4 address>|<IPv6 address>|<subnet>|<range>|<list>]
            [remoteip=any|localsubnet|dns|dhcp|wins|defaultgateway|<IPv4 address>|<IPv6 address>|<subnet>|<range>|<list>]
            [localport=0-65535|<port range>[,...]|RPC|RPC-EPMap|any]
            [remoteport=0-65535|<port range>[,...]|any]
            [protocol=0-255|icmpv4|icmpv6|icmpv4:type,code|icmpv6:type,code|tcp|udp|any]
        #>
        if ($fw.Rules.Count) {
            $direction_enum      = @{1 = 'in'; 2 = 'out'}
            $direction_enum_rev  = @{'in' = 1; 'out' = 2}
            $direction_name_enum = @{'Inbound' = 'in'; 'Outbound' = 'out'}

            if ($PSCmdlet.ParameterSetName -eq 'ByAttribute') {                
                $arguments = @()
                if ($All) {$arguments += ('name=all')}
                else      {$arguments += ('name={0}' -f $Name)}
                if ($PSBoundParameters.ContainsKey('Direction'))     {$arguments += ('dir={0}'        -f $direction_name_enum[$Direction] )}
		        if ($PSBoundParameters.ContainsKey('Profile'))       {$arguments += ('profile={0}'    -f ($Profile -join ',')             )}
		        if ($PSBoundParameters.ContainsKey('RemoteAddress')) {$arguments += ('remoteip={0}'   -f ($RemoteAddress -join ',')       )}
		        if ($PSBoundParameters.ContainsKey('Protocol'))      {$arguments += ('protocol={0}'   -f $Protocol                        )}
		        if ($PSBoundParameters.ContainsKey('LocalPort'))     {$arguments += ('localport={0}'  -f ($LocalPort -join ',')           )}
		        if ($PSBoundParameters.ContainsKey('RemotePort'))    {$arguments += ('remoteport={0}' -f ($RemotePort -join ',')          )}
		        if ($PSBoundParameters.ContainsKey('Program'))       {$arguments += ('program={0}'    -f $Program                         )}
		        if ($PSBoundParameters.ContainsKey('Service'))       {$arguments += ('service={0}'    -f $Service                         )}
            } else {
                $rule = $InputObject
                $arguments = @()
                $rule | gm -MemberType Property | Select -ExpandProperty Name | ? {$rule."$_"} | % {
                    $property_name = $_
                    switch ($property_name) {
                        Direction       {$arguments += ('dir={0}'        -f $direction_enum[$rule."$_"])}
                        Name            {$arguments += ('name={0}'       -f $rule.Name)}
                        RemoteAddresses {$arguments += ('remoteip={0}'   -f $rule."$_".Replace('*', 'any'))} 
                        RemotePorts     {$arguments += ('remoteport={0}' -f $rule."$_".Replace('*', 'any'))} 
                        ApplicationName {$arguments += ('program={0}'    -f $rule."$_")} 
                        serviceName     {$arguments += ('service={0}'    -f $rule."$_")}  
                        LocalPorts      {$arguments += ('localport={0}'  -f $rule."$_".Replace('*', 'any'))}
                        Profiles        { # convert mask to list of names
                                         $profile_enum = @{1 = 'domain'; 2 = 'private'; 4 = 'public'}
                                         $profile_mask = $rule."$_"
                                         $profile_list = ($profile_enum.Keys | ? {$profile_mask -band $_} | % {$profile_enum[$_]}) -join ','
                                         $arguments += ('profile={0}'    -f $profile_list)} 
                        Protocol        {$arguments += ('protocol={0}'   -f $rule."$_".ToString().Replace('256','any'))} 
                    }
                }
            }
            
            if ($PSCmdlet.ShouldProcess(($arguments -join ' '))) {
                $arguments = ('advfirewall', 'firewall', 'delete', 'rule') + $arguments
                Write-Debug ('Removing firewall rule with: netsh.exe ' + ($arguments -join ' '))
                $output = [string] (& netsh.exe $arguments 2>$1)
                if ($LASTEXITCODE -ne 0) {
                    if (-not $output.Contains('No rules match the specified criteria')) {
                        throw $output.Trim()
                    }
                }
            }
        }
	}
}
function Set-FirewallRule {
	[cmdletbinding(DefaultParameterSetName="ByName", SupportsShouldProcess=$True)] 
    param ( 
		[parameter(HelpMessage='The action to take for traffic which matches this rule.')]
		[ValidateSet('Allow', 'Block')] [string] 
            $Action,
		[parameter(HelpMessage='Information about the firewall rule.')] 
		[ValidateScript({-not $_.Contains('|')})] [string] 
            $Description,
		[parameter(HelpMessage='Which direction of traffic to match with this rule.')]
		[ValidateSet('Inbound','Outbound')] [string] 
            $Direction,
		[parameter(HelpMessage='Whether the rule is in effect or not.')] 
		[ValidateSet('True', 'False')] [string] 
            $Enabled,
        [string] 
            $Group,
        [ValidateScript( # Can be [0-255|*]:[0-255|*] | any
            {$_ -eq 'any' -or ((0..255 -contains $_.Split(':')[0]) -and (0..255+'*' -contains $_.Split(':')[1]))} 
        )] [string[]] 
            $IcmpType,
		[parameter(HelpMessage='Display name (alias) of the network interface that applies to the traffic.')] [string[]] 
            $InterfaceAlias,
		[parameter(HelpMessage='Only the specified network connection types are subject to this rule.')] 
		[ValidateSet('RemoteAccess','Wireless','Lan','All')] [string[]] 
            $InterfaceType,
        [ValidateScript( # Can be IPv[4|6][/cidr]|range
            {$_.Split('-') | % { [System.Net.IPAddress] ($_ -replace '/\d+$','') }}
        )] [string[]] 
            $LocalAddress,
		[parameter(HelpMessage='Network packets with matching IP port numbers match this rule.')]
        [ValidateScript( # Protocol must be TCP|UDP. Can be 1-65535-1-65535 | 1-65535
            {($_.Split('-') | ? {1..65535 -contains $_}).Count -eq $_.Split('-').Count}
        )] [string[]] 
            $LocalPort,
		[parameter(Mandatory=$true, ParameterSetName="ByName", 
            HelpMessage='Only the matching firewall rules of the indicated name are modified.')] 
		[ValidateScript({$_ -notmatch '^all$|\|'})] [string[]] 
            $Name,
		[parameter(HelpMessage='Change the name of the firewall rule.')] [ValidateScript({$_ -notmatch '^all$|\|'})] 
            [string] $NewName,
		[parameter(HelpMessage='Specifies one or more profiles to which the rule is assigned.')]
		[ValidateSet('Domain','Private','Public','Current','All')] [string[]] 
            $Profile,
		[parameter(HelpMessage='Specifies the path and file name of the program for which the rule allows traffic.')] [string] 
            $Program,
		[parameter(HelpMessage='Specifies that network packets with matching protocol match this rule.')]
        [ValidateScript( 
            {(0..255 -contains $_) -or ('Any','TCP','UDP','ICMPv4','ICMPv6' -contains $_)}
        )] [string] 
            $Protocol,
		[parameter(HelpMessage='Specifies that network packets with matching IP addresses match this rule.')]
        [ValidateScript( # Can be any|Defaultgateway|DHCP|WINS|LocalSubnet|IPv(4|6)addr or range|netmask or IPv4addr/cidr
            {((New-Object -ComObject HNetCfg.FWRule).RemoteAddresses = ($_ -replace 'any', '*')) -ne $null}
        )] [string[]] 
            $RemoteAddress,
        [parameter(HelpMessage='Network packets with matching IP port numbers match this rule.')]
        [ValidateScript( # Same as LocalPort
            {($_.Split('-') | ? {1..65535 -contains $_}).Count -eq $_.Split('-').Count}
        )] [string[]] 
            $RemotePort,
		[parameter(HelpMessage='Specifies the short name of a Windows service to which the firewall rule applies.')]
		[ValidateScript({((gsv|Select -Exp Name) + '*') -contains $_})] [string] 
            $Service,
		[parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true, HelpMessage='Firewall wall rule object to modify.')]
		    $InputObject
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
		$direction_enum = @{Inbound = 1; Outbound = 2}
        $protocol_enum = @{ICMPv4 = 1; ICMPv6 = 58; tcp = 6; udp = 17}

        $rule_names = $rules | Select -ExpandProperty Name
        $rule_names | Select -Unique | Sort | % {
            $rule_name = $_
            $count = ($rule_names | ? {$_ -eq $rule_name} | Measure).Count
            if ($count -gt 1) {$what_if_msg += 'name={0}({1}),' -f $rule_name, $count}
            else {$what_if_msg += 'name={0},' -f $rule_name}
        }

        if ($PSCmdlet.ShouldProcess($what_if_msg.TrimEnd(','))) {
		    if ($PSBoundParameters.ContainsKey('NewName')) {$rules | % {$_.Name = $NewName}}
		    if ($PSBoundParameters.ContainsKey('Description')) {$rules | % {$_.Description = $Description}}
		    if ($PSBoundParameters.ContainsKey('Action')) {$rules | % {$_.Action = $action_enum[$Action]}}
		    if ($PSBoundParameters.ContainsKey('Direction')) {$rules | % {$_.Direction = $direction_enum[$Direction]}}
		    if ($PSBoundParameters.ContainsKey('Enabled')) {$rules | % {$_.Enabled = ([bool]::Parse($Enabled))}}
		    if ($PSBoundParameters.ContainsKey('InterfaceAlias')) {
			    $nic_names = @(); $nic_names += $InterfaceAlias # Must be an array.
			    $rules | % {$_.Interfaces =$nic_names}}
		    if ($PSBoundParameters.ContainsKey('InterfaceType')) {
			    $interface_list = $InterfaceType -join ','
			    $rules | % {$_.InterfaceTypes = $interface_list}}
		    if ($PSBoundParameters.ContainsKey('Profile')) {
                $profile_mask = Get-FirewallProfileBitmask -Name $Profile
			    $rules | % {$_.Profiles = $profile_mask}}
		    if ($PSBoundParameters.ContainsKey('Program')) {$rules | % {$_.ApplicationName = $Program}}
	        if ($PSBoundParameters.ContainsKey('Protocol')) {$rules | % {
                    $i = 0
                    if ([int]::TryParse($Protocol, [ref] $i)) {$_.Protocol = $i}
                    else {$_.Protocol = $protocol_enum[$Protocol]}
                }
            }
            if ($PSBoundParameters.ContainsKey('IcmpType')) {$rules | % {$_.IcmpTypesAndCodes = $IcmpType.Replace('any','*')}}
		    if ($PSBoundParameters.ContainsKey('RemotePort')) {$rules | % {$_.RemotePorts = ($RemotePort -join ',')}}
		    if ($PSBoundParameters.ContainsKey('LocalPort')) {$rules | % {$_.LocalPorts = ($LocalPort -join ',')}}
            if ($PSBoundParameters.ContainsKey('LocalAddress')) {$rules | % {$_.LocalAddresses = ($LocalAddress -join ',')}}
            if ($PSBoundParameters.ContainsKey('RemoteAddress')) {$rules | % {$_.RemoteAddresses = ($RemoteAddress -join ',')}}
            if ($PSBoundParameters.ContainsKey('Service')) {$rules | % {$_.serviceName = $Service}}
            if ($PSBoundParameters.ContainsKey('Group')) {$rules | % {$_.Grouping = $Group}}
        }
	}
}
#endregion Rules ###########################################################################
#endregion Exports


#region Internal
function Get-FirewallConfigObject {
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa365309(v=vs.85).aspx
    return New-Object -ComObject HNetCfg.FwPolicy2
}

function Get-FirewallProfileBitmask {
    param (
        [parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public', 'current', 'all')] [string[]] $Name
    )
    $fw = Get-FirewallConfigObject
    $profile_enum = @{domain = 1; private = 2; public = 4; all = 2147483647; current = $fw.CurrentProfileTypes; selected = 0}
    $Name | % {$profile_enum.selected = $profile_enum.selected -bor $profile_enum[$_]}
    return $profile_enum.selected
}

function Set-FirewallProfileState {
    [cmdletbinding()]
    param ([int] $ProfileMask, [switch] $Enabled)
    $fw = Get-FirewallConfigObject
    Write-Verbose "Disabling firewall profile: $profile_name"
    $fw.FirewallEnabled($ProfileMask) = $Enabled
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
function Get-FirewallRule {
    # TODO support some filters
    (Get-FirewallConfigObject).Rules
}

function Get-FirewallProfile {
    # TODO create
}

function Set-FirewallProfile {
    [cmdletbinding(DefaultParameterSetName="ByName")]
    param (
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

function New-FirewallRule {
    [cmdletbinding()]
    param ( 
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateScript({$_ -notmatch '^all$|\|'})] [string] $Name, 
        [parameter(ParameterSetName="Make")] [ValidateScript({$_ -notmatch '\|'})] [string] $Description = $null,
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateScript({$_ -match '^\d{1,5}$|^\d{1,5}-\d{1,5}$'})] [string[]] $Port, # Supported Format: '80','443','8000-8009'
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('tcp', 'udp')] [string] $Protocol = 'tcp',
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('in', 'out')] [string] $Direction = 'in',
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('allow', 'block')] [string] $Action = 'allow',
        [parameter(ParameterSetName="Make")] [ValidateSet('domain', 'private', 'public', 'current', 'all')] [string[]] $Profile = 'all',
        [parameter(ParameterSetName="Make")] [ValidateSet("RemoteAccess", "Wireless", "Lan", "All")] [string[]] $InterfaceTypes = 'All',
        [parameter(ParameterSetName="Make")] [switch] $Disabled,
        [parameter(ParameterSetName="Make")] [string] $GroupingName = '',
        [parameter(ParameterSetName="Make")] [string] $ApplicationPath = '',
        [parameter(ParameterSetName="Make")] [ValidateScript({ ((gsv | Select -Exp Name) + '*') -contains $_})] [string] $ServiceShortName = '',
        [parameter(ParameterSetName="Make")] [string[]] $InterfaceNames = '',
        [parameter(ParameterSetName="Made")] $FirewallRuleObject
    )

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
        try {
            $fw.Rules.Add($rule)
            return $rule
        } catch {
            $PSCmdlet.WriteError($_)
        }
    }
}

function Remove-FirewallRule {
    [cmdletbinding()]
    param ([parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] [string[]] $Name)
    process {
        # TODO accept rule object and re-add all rules with the same name except the one matching all properties
        # TODO support wildcards
        # TODO support -Confirm and -Whatif
        $fw = Get-FirewallConfigObject
        foreach ($rule_name in $Name) {
            $fw.Rules | ? {$_.Name -eq $rule_name} | % {
                $rule = $_
                $rule_print = ($rule | gm -MemberType Property | % {"{0} : {1}" -f $_.Name, $rule."$($_.Name)"}) -join "`n"
                Write-Verbose ("Removing firewall rule:`n" + $rule_print)
                # TODO warn when multiple rules matched
                $fw.Rules.Remove($rule_name)
            }
        }
    }
}

function Set-FirewallRule {
    [cmdletbinding()]
    param ( 
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateScript({$_ -notmatch '^all$|\|'})] [string] $Name, 
        [parameter(ParameterSetName="Make")] [ValidateScript({-not $_.Contains('|')})] [string] $Description = $null,
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateScript({$_ -match '^\d{1,5}$|^\d{1,5}-\d{1,5}$'})] [string[]] $Port, # Supported Format: '80','443','8000-8009'
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('tcp', 'udp')] [string] $Protocol,
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('in', 'out')] [string] $Direction,
        [parameter(ParameterSetName="Make", Mandatory=$true)] [ValidateSet('allow', 'block')] [string] $Action,
        [parameter(ParameterSetName="Make")] [ValidateSet('domain', 'private', 'public', 'current', 'all')] [string[]] $Profile = 'all',
        [parameter(ParameterSetName="Make")] [ValidateSet("RemoteAccess", "Wireless", "Lan", "All")] [string[]] $InterfaceTypes = 'All',
        [parameter(ParameterSetName="Make")] [switch] $Disabled,
        [parameter(ParameterSetName="Make")] [string] $GroupingName = '',
        [parameter(ParameterSetName="Make")] [string] $ApplicationPath = '',
        [parameter(ParameterSetName="Make")] [ValidateScript({ ((gsv | Select -Exp Name) + '*') -contains $_})] [string] $ServiceShortName = '',
        [parameter(ParameterSetName="Make")] [string[]] $InterfaceNames = '',
        [parameter(ParameterSetName="Made")] $FirewallRuleObject
    )
    
    # TODO accept objects from the pipeline
    # TODO warn when name matches more than one rule
    # TODO see if an item in the rule collection can be directly modified: $fw.rules.item(?)
    # TODO check about getting the rule uniquely based on HashCode
    # TODO combine Disable|Enable-FirewallRule here
}

function Disable-FirewallRule {
    [cmdletbinding()]
    param ([parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] [string[]] $Name)
    process {
        $fw = Get-FirewallConfigObject
        foreach ($rule_name in $Name) {
            $fw.Rules | ? {$_.Name -eq $rule_name} | % {
                $rule = $_
                $rule_print = ($rule | gm -MemberType Property | % {"{0} : {1}" -f $_.Name, $rule."$($_.Name)"}) -join "`n"
                Write-Verbose ("Disabling firewall rule:`n" + $rule_print)
                $rule.Enabled = $false
                Remove-FirewallRule -Name $rule.Name
                New-FirewallRule -FirewallRuleObject $rule
            }
        }
    }
}

function Enable-FirewallRule {
    [cmdletbinding()]
    param ([parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] [string[]] $Name)
    process {
        $fw = Get-FirewallConfigObject
        foreach ($rule_name in $Name) {
            $fw.Rules | ? {$_.Name -eq $rule_name} | % {
                $rule = $_
                $rule_print = ($rule | gm -MemberType Property | % {"{0} : {1}" -f $_.Name, $rule."$($_.Name)"}) -join "`n"
                Write-Verbose ("Disabling firewall rule:`n" + $rule_print)
                $rule.Enabled = $true
                Remove-FirewallRule -Name $rule.Name
                New-FirewallRule -FirewallRuleObject $rule
            }
        }
    }
}
#endregion

Export-ModuleMember -Function Get-FirewallRule, 
        Get-FirewallProfile, Set-FirewallProfile,
        New-FirewallRule, Remove-FirewallRule, Set-FirewallRule, Disable-FirewallRule, Enable-FirewallRule

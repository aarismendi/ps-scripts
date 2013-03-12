#region Internal
function Get-FirewallConfigObject {
    # http://msdn.microsoft.com/en-us/library/windows/desktop/aa365309(v=vs.85).aspx
    return New-Object -ComObject HNetCfg.FwPolicy2
}

function Get-FirewallProfileBitmask {
    param (
        [parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public', 'current', 'all')] [string[]] $Profile
    )
    $fw = Get-FirewallConfigObject
    $profile_enum = @{domain = 1; private = 2; public = 4; all = 2147483647; current = $fw.CurrentProfileTypes; selected = 0}
    $Profile | % {$profile_enum.selected = $profile_enum.selected -bor $profile_enum[$_]}
    return $profile_enum.selected
}

function Set-FirewallProfileState {
    [cmdletbinding()]
    param ([int] $ProfileMask, [switch] $Enabled)
    $fw = Get-FirewallConfigObject
    Write-Verbose "Disabling firewall profile: $profile_item"
    $fw.FirewallEnabled($ProfileMask) = $Enabled
}
#endregion

#region Exports
function Get-FirewallRule {
    # TODO support some filters
    (Get-FirewallConfigObject).Rules
}

function Disable-FirewallProfile {
    [cmdletbinding()]
    param ([parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Profile)
    foreach ($profile_item in $Profile) {
        $profile_mask = Get-FirewallProfileBitmask -Profile $profile_item
        Set-FirewallProfileState -ProfileMask $profile_mask -Enabled:$false
    }
}

function Enable-FirewallProfile {
    [cmdletbinding()]
    param ([parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Profile)
    foreach ($profile_item in $Profile) {
        $profile_mask = Get-FirewallProfileBitmask -Profile $profile_item
        Set-FirewallProfileState -ProfileMask $profile_mask -Enabled
    }
}

function Set-InboundTrafficAction {
    [cmdletbinding()]
    param (
        [parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Profile,
        [parameter(Mandatory=$true)] [ValidateSet('allow', 'block', 'block_all', 'default')] [string] $Action
    )
    $fw = Get-FirewallConfigObject
    $action_enum = @{block = 0; allow = 1; default = 0}
    foreach ($profile_item in $Profile) {
        $profile_mask = Get-FirewallProfileBitmask -Profile $profile_item
        if ($Action -eq 'block_all') {
            $fw.BlockAllInboundTraffic($profile_mask) = $true
        } else {
            $fw.BlockAllInboundTraffic($profile_mask) = $false
            $fw.DefaultInboundAction($profile_mask) = $action_enum[$Action]
        }
    }
}

function Set-OutboundTrafficAction {
    [cmdletbinding()]
    param (
        [parameter(Mandatory=$true)] [ValidateSet('domain', 'private', 'public')] [string[]] $Profile,
        [parameter(Mandatory=$true)] [ValidateSet('allow', 'block', 'default')] [string] $Action
    )
    $fw = Get-FirewallConfigObject
    $action_enum = @{block = 0; allow = 1; default = 1}
    foreach ($profile_item in $Profile) {
        $profile_mask = Get-FirewallProfileBitmask -Profile $profile_item
        $fw.DefaultOutboundAction($profile_mask) = $action_enum[$Action]
    }
}

function Set-FirewallProfile {
    param ()
    # TODO combine Set-InboundTrafficAction, Set-OutboundTrafficAction, Enable-FirewallProfile, Disable-FilewallProfile
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
                $profile_mask = Get-FirewallProfileBitmask -Profile $Profile

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
        [parameter(ParameterSetName="Make")] [ValidateScript({-not ($_.Contains('|')})] [string] $Description = $null,
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
        Disable-FirewallProfile, Enable-FirewallProfile,
        Set-InboundTrafficAction, Set-OutboundTrafficAction, 
        New-FirewallRule, Remove-FirewallRule, Disable-FirewallRule, Enable-FirewallRule

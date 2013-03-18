cls
# Reset environment
Remove-Variable * -ErrorAction SilentlyContinue

$ErrorActionPreference = 'Stop'
$this_path = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path

Import-Module -Name "$this_path\WinFirewall.psm1" -Force -Verbose

Write-Host "Backing up existing firewall config"
$wfw_file_name = [IO.Path]::ChangeExtension([guid]::NewGuid().Guid, 'wfw')
$wfw_file_path = Join-Path -Path $env:TEMP -ChildPath $wfw_file_name
& netsh.exe advfirewall export "$wfw_file_path"

try {
Remove-FirewallRule -All -Verbose
Write-Host "Running New-FirewallRule Tests" -ForegroundColor Magenta
. "$this_path\Test\New-FirewallRule.Test.ps1"

Write-Host "Running Set-FirewallRule Tests" -ForegroundColor Magenta
. "$this_path\Test\Set-FirewallRule.Test.ps1"

Write-Host "Running Remove-FirewallRule Tests" -ForegroundColor Magenta
. "$this_path\Test\Remove-FirewallRule.Test.ps1"
# TODO . "$this_path\Test\Get-FirewallRule.Test.ps1"
# TODO . "$this_path\Test\Get-FirewallProfile.Test.ps1"
# TODO . "$this_path\Test\Set-FirewallProfile.Test.ps1"
} catch {
    $_
} finally {
    Write-Host "Restoring original firewall config."
    if (Test-Path -Path $wfw_file_path -PathType Leaf) {
        & netsh.exe advfirewall import "$wfw_file_path"
        Remove-Item -Path $wfw_file_path -Force
    }
}
<#
The scripts is to be run every few minutes. Its purpose it to move
computers (non-server) to a more agreeable OU so that GPO's can be applied
without extra effort.
#>
[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)][PSCredential]$ADCredential,
 [Parameter( Mandatory = $True)][string]$TargetOrgUnitPath,
 [Parameter(Mandatory = $False)][string[]]$ExcludedOUs,
 [Alias('wi')][switch]$WhatIf
)
function Get-StaleComputers ( [pscredential]$cred) {
 $props = 'Name', 'DistinguishedName', 'ObjectGUID', 'LastLogonDate', 'OperatingSystem', 'Enabled'
 $cutoffDate = (Get-Date 12:00AM).AddYears(-1)
 $filter = "LastLogonDate -lt `$cutoffDate -and -not(OperatingSystem -like '*Server*')"
 $params = @{
  Filter     = $filter
  Properties = $props
  Credential = $cred
 }
 $result = Get-ADComputer @params | Where-Object {
  $_.DistinguishedName -notlike "*$ExcludedOUs*" -and
  $_.DistinguishedName -notlike '*Server*' -and
  $_.OperatingSystem -notlike '*Server*'
 } | Select-Object -Property $props
 Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, @($result).Count)
 $result
}

function Disable-Stale ( [pscredential]$cred) {
 process {
  $oldOU = $_.DistinguishedName -ireplace "CN=$($_.Name),", ''
  $desc = "Disabled by Jenkins:$(Get-Date -f 'yyyy-MM-dd') OldOU: $oldOU"
  $params = @{
   Identity    = $_.ObjectGUID
   Enabled     = $false
   Description = $desc
   Credential  = $cred
   WhatIf      = $WhatIf
  }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.name, $desc) -F Darkblue
  Set-ADComputer @params
  $_
 }
}

function Move-Stale ( $ou, [pscredential]$cred) {
 process {
  $params = @{
   Identity   = $_.ObjectGUID
   TargetPath = $ou
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.name, $ou) -F Blue
  Move-ADObject @params
  $_
 }
}

function Remove-Stale ( [pscredential]$cred) {
 process {
  $params = @{
   Identity   = $_.ObjectGUID
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.name) -F Red
  Remove-ADObject @params
  $_
 }
}

function Select-Enabled {
 process {
  $_ | Where-Object { $_.Enabled -eq $true }
 }
}

function Select-OldStale {
 process {
  $_ | Where-Object { $_.Enabled -eq $false -and $_.LastLogonDate -lt (Get-Date).AddYears(-2) }
 }
}

function Show-Object {
 process {
  Write-Verbose ($_ | Out-String)
 }
}

# ===============================================================================================================
Import-Module CommonScriptFunctions
Import-Module -Name ActiveDirectory -Cmdlet Get-ADComputer, Set-ADComputer, Move-ADObject, Remove-ADObject -WarningAction SilentlyContinue -ErrorAction Stop

if ($WhatIf) { Show-TestRun }

$staleComputers = Get-StaleComputers -cred $ADCredential
$staleComputers |
 Select-Enabled |
  Disable-Stale -cred $ADCredential |
   Move-Stale -ou $TargetOrgUnitPath -cred $ADCredential |
    Show-Object
$staleComputers |
 Select-OldStale |
  Remove-Stale -cred $ADCredential |
   Show-Object

if ($WhatIf) { Show-TestRun }
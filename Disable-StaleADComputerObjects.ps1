<#
The scripts is to be run every few minutes. Its purpose it to move
computers (non-server) to a more agreeable OU so that GPO's can be applied
without extra effort.
#>
[cmdletbinding()]
param (
 [Parameter(Position = 0, Mandatory = $True)]
 [Alias('DC', 'Server')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 1 })]
 [string]$DomainController,
 [Parameter(Position = 1, Mandatory = $True)]
 [Alias('ADCred', 'ADCReds')]
 [System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Position = 2, Mandatory = $True)]
 [Alias('SrcOU')]
 [string]$SourceOrgUnitPath,
 [Parameter(Position = 3, Mandatory = $True)]
 [Alias('TargOU')]
 [string]$TargetOrgUnitPath,
 [switch]$WhatIf
)

Get-PSSession | Remove-PSSession
# AD Domain Controller Session
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -AllowClobber

$cutoff = (Get-date).addyears(-1)

$params = @{
 Properties = 'lastLogonTimestamp'
 SearchBase = $SourceOrgUnitPath
}

$staleComputerObjs = Get-ADComputer @params -Filter * | Where-Object { [datetime]::FromFileTime($_.lastLogonTimestamp) -lt $cutoff }
Write-Host 'Stale Computer Objects:'

$staleComputerObjs | ForEach-Object {
 $desc = "Disabled by Jenkins on $(Get-Date -f 'yyyy-MM-dd')"
 Write-Host ('[{0}] Disabling and moving stale object' -f $_.name)
 Set-ADComputer -Identity $_.ObjectGUID -Enabled $false -Description $desc -WhatIf:$WhatIf
 Move-ADObject -Identity $_.ObjectGUID -TargetPath $TargetOrgUnitPath -WhatIf:$WhatIf
}

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession
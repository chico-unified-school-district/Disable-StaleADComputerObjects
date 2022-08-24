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
 [Parameter(Position = 4, Mandatory = $False)]
 [Alias('ExOU')]
 [string[]]$ExcludedOUs,
 [switch]$WhatIf
)

Get-PSSession | Remove-PSSession
# AD Domain Controller Session
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
$cmdlets = 'Get-ADComputer', 'Set-ADComputer', 'Move-ADObject'
Import-PSSession -Session $adSession -CommandName $cmdlets -Module ActiveDirectory -AllowClobber | Out-Null

$cutoff = (Get-date).addyears(-1)

$params = @{
 Filter     = { Enabled -eq $True }
 Properties = 'lastLogonTimestamp'
 SearchBase = $SourceOrgUnitPath
}

$staleComputerObjs = Get-ADComputer @params | Where-Object {
 [datetime]::FromFileTime($_.lastLogonTimestamp) -lt $cutoff -and
 $_.Enabled -eq $True
}

Write-Host ('Stale Computer Objects: {0}' -f ($staleComputerObjs | Measure-Object).count)

$staleComputerObjs | ForEach-Object {
 $oldOu = $_.DistinguishedName
 $desc = "Disabled by Jenkins:$(Get-Date -f 'yyyy-MM-dd') OldOU:$oldOU"
 Write-Host ('[{0}] Disabling stale object' -f $_.name)
 Set-ADComputer -Identity $_.ObjectGUID -Enabled $false -Description $desc -WhatIf:$WhatIf
 foreach ($ou in $ExcludedOUs) {
  if ($_.DistinguishedName -like "*$ou*") {
   Write-Verbose "$($_.name),Excluded OU. Skipping this computer object"
   continue
  }
  Write-Host ('[{0}] Moving stale object' -f $_.name)
  Move-ADObject -Identity $_.ObjectGUID -TargetPath $TargetOrgUnitPath -WhatIf:$WhatIf
 }
}

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession
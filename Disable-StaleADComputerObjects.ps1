<#
The scripts is to be run every few minutes. Its purpose it to move
computers (non-server) to a more agreeable OU so that GPO's can be applied
without extra effort.
#>
[cmdletbinding()]
param (
 [Parameter(Position = 0, Mandatory = $True)]
 [Alias('DCs')]
 [string[]]$DomainControllers,
 [Parameter(Position = 1, Mandatory = $True)]
 [Alias('ADCred')]
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
 [Alias('wi')]
 [switch]$WhatIf
)

. .\lib\New-ADSession.ps1
. .\lib\Select-DomainController.ps1
. .\lib\Show-TestRun.ps1

Show-TestRun

Get-PSSession | Remove-PSSession

$dc = Select-DomainController $DomainControllers
$cmdlets = 'Get-ADComputer', 'Set-ADComputer', 'Move-ADObject', 'Remove-ADObject'
New-ADSession -dc $dc -cmdlets $cmdlets -cred $ADCredential

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
 # $oldOu = $_.DistinguishedName
 $oldOu = $_.DistinguishedName -ireplace "CN=$($_.Name),", ''
 $desc = "Disabled by Jenkins:$(Get-Date -f 'yyyy-MM-dd') OldOU: $oldOU"
 Write-Host ('[{0}] Disabling stale object,[{1}]' -f $_.name, $desc) -Fore Blue
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

$DeadParams = @{
 Filter     = { Enabled -eq $False }
 Properties = 'lastLogonTimestamp'
 SearchBase = $TargetOrgUnitPath
}

$DeadComputerObjects = Get-ADComputer @DeadParams |
Where-Object {
 ([datetime]::FromFileTime($_.lastLogonTimestamp) -lt (Get-Date).AddYears(-2)) -and
 ($_.DistinguishedName -notlike '*servers*')
} |
Sort-Object LastLogonTimeStamp

$DeadComputerObjects |
ForEach-Object {
 $LastLogonDate = [datetime]::FromFileTime($_.lastLogonTimestamp)
 Write-Host ('Deleting {0} {1}' -F $_.Name , $LastLogonDate) -F Cyan
 Remove-ADObject -Identity $_.ObjectGUID -Recursive -Confirm:$False -WhatIf:$WhatIf
}

Write-Host ('Dead Computer Objects: {0}' -f ($DeadComputerObjects | Measure-Object).count)

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession
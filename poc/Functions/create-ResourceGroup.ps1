<#
.SYNOPSIS
  Creates the resource group specified in the parameter

.PARAMETER name
  Name of resource group to be removed

  
.PARAMETER tags
  A hash table of key value pairs

.PARAMETER siteLocation
  Azure region


.EXAMPLE
  .\create-rg.ps1 -name "hiscoxrg01"

#>

Param (
   [Parameter(Mandatory=$true)]
   [string]$name,
   [Parameter(Mandatory=$true)]
   [hashtable]$tags,
   [Parameter(Mandatory=$true)]
   [string]$siteLocation
)

If (Get-AzureRmResourceGroup | Where-Object {$_.ResourceGroupName -eq $name}) {
  Throw New-Object System.Exception("Resource group found")
}

New-AzureRmResourceGroup -Name $name -Location $siteLocation -Tag $tags

$timeout = New-TimeSpan -Minutes 1
$sw = [Diagnostics.Stopwatch]::StartNew()

Do {
  If (Get-AzureRmResourceGroup | Where-Object {$_.ResourceGroupName -eq $name -and $_.ProvisioningState -eq "Succeeded"}) {
    Return
  }
} While ($sw.Elapsed -lt $timeout)

Throw New-Object System.Exception("Resource group not found")

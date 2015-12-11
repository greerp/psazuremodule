<#
.SYNOPSIS
  Removes the resource group

.PARAMETER name
  Name of resource group to be removed

.EXAMPLE
  .\remove-resourcegroup.ps1 -name "hiscoxrg01"

#>

Param (
  [Parameter(Mandatory=$true)]
  [string]$name
)

Get-AzureRmResourceGroup | Where-Object {$_.ResourceGroupName -eq $name} | Remove-AzureRmResourceGroup -Force

$timeout = New-TimeSpan -Minutes 1
$sw = [Diagnostics.Stopwatch]::StartNew()

Do {
  If (-not (Get-AzureRmResourceGroup | Where-Object {$_.ResourceGroupName -eq $name})) {
    Return
  }
} While ($sw.Elapsed -lt $timeout)

Throw New-Object System.Exception("Resource group still exists!!!")

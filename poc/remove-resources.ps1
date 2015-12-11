<#
.SYNOPSIS
  Removes all resources from Azure resource group and deletes the resource group

.EXAMPLE
  .\remove-resources.ps1 -resourcegroupname "helloworld_RG_01_0.0.1"

#>
Param (
  [string]$resourcegroupname=[string]::Empty
)

# our function scripts
$azureAuthenticate = [System.IO.Path]::Combine($PSScriptRoot, "Functions\azure-authenticate.ps1")
$removeResourceGroup = [System.IO.Path]::Combine($PSScriptRoot, "Functions\remove-resourcegroup.ps1")

Try {
  If ([string]::IsNullOrEmpty($resourcegroupname)) {
    Throw New-Object System.Exception("Resource group name is not set")
  }
  & $azureAuthenticate
  & $removeResourceGroup -name $resourcegroupname
} Catch {
  Write-Host "An error occured in   : $($_.InvocationInfo.ScriptName)"
  Write-Host "The error was at line : $($_.InvocationInfo.ScriptLineNumber)"
  Write-Host "The actual error is   : $($_.Exception.Message)"
  Exit 1
}

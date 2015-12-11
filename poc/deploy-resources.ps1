<#
.SYNOPSIS
  Create a resource group and deploy all resources into Azure

.PARAMETER resourceGroupName
  Name of resource group to be created

.PARAMETER deploymentGroupName
  Name of deployment group to be created

.PARAMETER siteLocation
    Azure region
    depl
.PARAMETER versionNumber
  Version number

.EXAMPLE
  .\deploy-resources.ps1 -resourceGroupName "helloworld_RG_01_0.0.1"

#>

Param (
  [string]$resourceGroupName=[string]::Empty,
  [string]$deploymentGroupName=[string]::Empty,
  [string]$siteLocation=[string]::Empty,
  [string]$versionNumber=[string]::Empty
)
$templateFile=".\resourcestemplate.json"

$tags=@{name="project";value="azure-poc"}
$versionNumber= $versionNumber -replace "\.", "-"
$templateParameters=@{resourceSuffix=$versionNumber}

Try {
  If ([string]::IsNullOrEmpty($resourceGroupName)) {
    Throw New-Object System.Exception("Resource group name is not set")
  }
  If ([string]::IsNullOrEmpty($deploymentGroupName)) {
    Throw New-Object System.Exception("Deployment group name is not set")
  }
  If ([string]::IsNullOrEmpty($siteLocation)) {
    Throw New-Object System.Exception("Site Location is not set")
  }
  . .\Functions\azure-authenticate.ps1
  . .\Functions\create-resourcegroup.ps1 -name $resourceGroupName -tags $tags -siteLocation $siteLocation
  . .\Functions\create-resources.ps1 -Name $deploymentGroupName -ResourceGroupName $resourceGroupName -TemplateFile $templateFile -TemplateParameters $templateParameters
} Catch {
  Write-Host "An error occured in   : $($_.InvocationInfo.ScriptName)"
  Write-Host "The error was at line : $($_.InvocationInfo.ScriptLineNumber)"
  Write-Host "The actual error is   : $($_.Exception.Message)"
  Exit 1
}

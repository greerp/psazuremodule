<#
.SYNOPSIS
  Creates the resources in specified resource group

.PARAMETER name
  Name of the deployment

.PARAMETER resourceGroupName
  Name of the resource group to deploy the resources to

.PARAMETER templateFile
  Path to the JSON template file

.PARAMETER templateParameters
  Hash table of template parameters

.EXAMPLE
  .\create-rg.ps1 -name "hiscoxrg01"

#>

Param (
   [Parameter(Mandatory=$true)]
   [string]$name,
   [Parameter(Mandatory=$true)]
   [string]$resourceGroupName,
   [Parameter(Mandatory=$true)]
   [string]$templateFile,
   [Parameter(Mandatory=$true)]
   [hashtable]$templateParameters

)

New-AzureRmResourceGroupDeployment -Name $name -ResourceGroupName $resourceGroupName -TemplateFile $templateFile -TemplateParameterObject $templateParameters

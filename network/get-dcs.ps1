<#
.SYNOPSIS
  Lists all the DC on the specified domain 

.PARAMETER domainName
  Name of resource group to be created


.EXAMPLE
  .\get-alldcs.ps1 -domainName "insidepixie"

#>

Param (
    [Parameter(Mandatory=$true )] [string] $domainName
)


$type = [System.DirectoryServices.ActiveDirectory.DirectoryContextType]"Domain"
$context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext($type, $domainName)
$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)

ForEach ($dc in $domain.FindAllDomainControllers() ) { 
    $dc.Name 
}


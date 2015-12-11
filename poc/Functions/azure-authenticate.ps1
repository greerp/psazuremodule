 <#
.SYNOPSIS
  Authenticates Azure service account

.EXAMPLE
  '.\azure-authenticate.ps1'

#>

$encryptedFile = [System.IO.Path]::Combine($PSScriptRoot, "encryptedcred.xml")
$username="SVC_Nimbus@hiscox.com"

# Read xml from bitbucket repository
$object = Import-Clixml -Path $encryptedFile

# Retrive the certificate from the vault using the thumbprint
$thumbprint = '74061A0EC377BBA47A4F9DDE8072BF7F3228C22D'
$cert = Get-Item -Path Cert:\localmachine\My\$thumbprint -ErrorAction Stop

# Decrypt the encrypted key using the cert private key
$key = $cert.PrivateKey.Decrypt($object.Key, $true)

# Decrypt the password using the decrypted key form previous line
# $secureKey now contain the service account password as a securestring
$secureString = $object.Payload | ConvertTo-SecureString -Key $key

$cred = New-Object System.Management.Automation.PSCredential ($userName, $secureString)

Add-AzureRMAccount -credential $cred

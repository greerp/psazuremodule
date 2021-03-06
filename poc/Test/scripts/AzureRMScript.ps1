﻿"C:\Program Files (x86)\Microsoft SDKs\Azure\PowerShell\ServiceManagement\Azure\Services\ShortcutStartup.ps1"

$encryptedFile = [System.IO.Path]::Combine($PSScriptRoot, "encryptedcred.xml")
#$encryptedFile = "C:\encryptedcred.xml"
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

# Add the RM account to current context
#Add-AzureAccount -Credential $cred
Add-AzureRMAccount -Credential $cred

# Automated tests will append commands below:
Get-AzureSubscription
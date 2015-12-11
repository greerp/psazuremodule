Param(
  [string]$accountname
)

if ( $accountname -eq $null ) {

    $serviceAccount = Read-Host -Prompt "Enter Service Account Name: "
}
else {
    $serviceAccount = $accountname
}

$objectFile =  ".\account-creds.xml"

# Read the Encrypted Key and credentials file into a PS object. 
# The Object was previously serialised using whaver format PS uses for object serialisation
# The original object was a anonymous object with two properties: key and payload
$object = Import-Clixml $objectFile


# Read the associated cert from the localmachine vault
$cert = Get-Item -Path Cert:\LocalMachine\My\4EC0E3E7F31F26E27F3ED9F444ED9F91DF412309 -ErrorAction Stop

# Using the cert private key, decrypt the key stored in the credentials object
$key = $cert.PrivateKey.Decrypt($object.Key, $true)

# Decrypt the password using the decrypted key and save in SecureString object
$securePassword = $object.PayLoad | ConvertTo-SecureString -Key $key

$credential = New-Object System.Management.Automation.PSCredential ($serviceAccount, $securePassword)

Add-AzureAccount -Credential $credential



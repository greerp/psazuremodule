   <#
.SYNOPSIS
Creates encrypted password using the certificate and stores within an XML file



.EXAMPLE
   '.\created encrypted password.ps1' 

#>
$encryptedFile = ".\encryptedcred.xml"
   
   
$secureString = Read-Host -AsSecureString -Prompt "Enter Password"
   
# We convert the AD Service Account Password into a secure string which means i can not be read in clear text programatically.
    
#$secureString = 'Azure Service Account Password' | ConvertTo-SecureString -AsPlainText -Force

# Generate our new 32-byte AES key.  I don't recommend using Get-Random for this; the System.Security.Cryptography namespace
# offers a much more secure random number generator.

# The Key is used to encrypt the AD service account password.
# Use crypto random number generate to create a unique key and allocate it to a 32 byte array
$key = New-Object byte[](32)
[System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($key)


# This command encrypts the password with the key.
$encryptedString = ConvertFrom-SecureString -SecureString $secureString -Key $key

# This is the thumbprint of a certificate on my test system where I have the private key installed.
# The thumbprint must be read from the certificate to identify it wthin powershell.

$thumbprint = '74061A0EC377BBA47A4F9DDE8072BF7F3228C22D'
#Retrieve the certificate using the thumbprint
$cert = Get-Item -Path Cert:\LocalMachine\My\$thumbprint -ErrorAction Stop

# We use the public key on the cert to encrypt the key
$encryptedKey = $cert.PublicKey.Key.Encrypt($key, $true)

    
# Create an PS dynamic object containing the key encrypted with he cert pk and the password that has been 
# encrypted using the unencrypted key
$object = New-Object psobject -Property @{
    Key = $encryptedKey
    Payload = $encryptedString
}

# Save the dynamic object to an XML file for storage on public repository such as bitbucket
$object | Export-Clixml $encryptedFile





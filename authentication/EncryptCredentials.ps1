# A certificate needs to be generated and stored in the local machine vault under personal certificates
# You can use IIS manager to issue an example cert of use one issued by a CA


# Convert the service Account Password to a Secure String. 
# A secure String is just a variable type that cannot be read programmtically
$securePassword = Read-Host -AsSecureString -Prompt "Enter Password"
$objectFile =  ".\account-creds.xml"

#$securePassword = "<Password>" | ConvertTo-SecureString -AsPlainText -Force

# Generate a key that will be used to encrypt our password, store it in a 32 byte array
$key = New-Object byte[](32)
[System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($key)

# Now encrypt the password using the randomly generated key
$encryptedPassword = ConvertFrom-SecureString -SecureString $securePassword -Key $key

# Retrieve the certificate using the certificates thumbprint, this can be opbtained by looking in the 
# vault at the certificates properties
$cert = Get-Item -Path Cert:\LocalMachine\My\4EC0E3E7F31F26E27F3ED9F444ED9F91DF412309 -ErrorAction Stop


# Use the public key of the certificate to encrypt our key that we used to encrypt the password
# When we come to decrypt the password we use the certificate's private key to decrypt $encryptedKey we 
# can then use key to decrypt the password 
$encryptedKey = $cert.PublicKey.Key.Encrypt($key,$true)

# Create a PS dynamic object that contains the encrypted key and the encrypted password. So you need this file 
# along with the certificate to obtain the service account cerdentials

$object = New-Object psobject -Property @{
    Key = $encryptedKey
    Payload = $encryptedPassword
}

Write-Host "Encrypted Credentials stored in $objectFile"
$object | Export-Clixml $objectFile



#http://site242.azurewebsites.net/

#New-AzureADApplication -displayname "pgsecapp1" -HomePage "http://site242.azurewebsites.net/" -IdentifierUris "http://site242.azurewebsites.net/"



$subscriptionId = "311818f8-d369-419b-bfe1-fdf644de096f"; 
$aadTenantDomain = "pgdirectory.onmicrosoft.com"; 
$aadClientId = "f13751a2-45ab-4106-8314-711f2e4438d7"; 

$authString = "https://login.windows.net/$aadTenantDomain"

# Creates a context for login.windows.net (Azure AD common authentication)
[Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext]$authContext = 
    [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext]$authString

[Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential] $cred = 
    New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential "service@${aadTenantDomain}", "R3dpixie"


$authContext.AcquireToken("https://management.core.windows.net/", 
        $aadClientId, 
        $cred)



<#

private const string _subscriptionId = "xxxxxxxxxxxxxxxxx"; 
private const string _aadTenantDomain = "tomhollanderhotmail.onmicrosoft.com"; 
private const string _aadClientId = "yyyyyyyyyyyyy"; 

private static string GetAuthorizationHeader() 
{ 
    AuthenticationResult result = null; 
    var context = new AuthenticationContext("https://login.windows.net/" + _aadTenantDomain);

 

    // If you wanted to show a credential dialog, do this: 
    //result = context.AcquireToken( 
    //    "https://management.core.windows.net/", 
    //    _aadClientId, 
    //      new Uri("http://localhost"), PromptBehavior.Auto);

    // Directly specify the username and password. 
    var credential = new UserCredential( 
        ConfigurationManager.AppSettings["serviceAccountUserName"], 
        ConfigurationManager.AppSettings["serviceAccountPassword"]); 
    result = context.AcquireToken( 
        "https://management.core.windows.net/", 
        _aadClientId, 
            credential); 
    if (result == null) 
    { 
        throw new InvalidOperationException("Failed to obtain the JWT token"); 
    }

    #>
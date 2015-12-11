<#
    .SYNOPSIS
      Allocate CIDR from SQL subnet database

    .DESCRIPTION
      Attempts to allocate adn return a CIDR from the database,
      f -verbose is specified error messages will be writttn to stdout. If not errors will result in an exception.
      If $user is not specified, attempt to will be made via trusted connection   

    .PARAMETER resourceGroupName
      Name of resource group to be created

    .PARAMETER range
      Minimum requirment of hosts in subnet 

    .PARAMETER server
      Sql Server 
        
    .PARAMETER database 
      Sql Database

    .PARAMETER user 
      Database user

    .PARAMETER password
      Database password 

    .EXAMPLE
      . .\get-cidr.ps1 -resourceGroup home -range 10 -server pghiscoxsqlsrv.database.windows.net -sqlDatabase pghiscox 

#>
Param(
    [Parameter(Mandatory=$true)][string]$resourceGroup,
    [Parameter(Mandatory=$true)][int]$range,
    [string]$server,
    [string]$database,
    [string]$user,
    [string]$password
)

    $verbose= $PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent


    if ($verbose -eq $true ){
        Write-Host "Using SQL Server: ${server}, database: ${database}"
    }

    $sqlConnection = New-Object System.Data.SQLClient.SQLConnection


    if ([string]::IsNullOrEmpty($user)){
        $sqlConnection.ConnectionString = 
	        "server=${server};database=${database};trusted_connection=true;"
    }
    else {
        $sqlConnection.ConnectionString = 
            "server=${server};database=${database};user ID=${user};Password=${password};Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"
    }
		
    try {
	    $sqlConnection.Open()
    }
    catch {
        if ($verbose) {
  	        $ErrorMessage = $_.Exception.Message
  	        Write-Host $ErrorMessage
    	    Exit
        }
        else {
            throw 
        }
    }

    $sqlCommand = New-Object System.Data.SQLClient.SQLCommand
    $sqlCommand.Connection = $sqlConnection
	$sqlCommand.CommandText = "exec pr_getresgroupcidr @resgroup='${resourceGroup}', @range=${range}"
	try {
		$cidr = $sqlCommand.ExecuteScalar()
        write-host $cidr.ToString();
		}
	catch {
        $err = $_.Exception
        if ($verbose) {
            Write-Host $err.Message
            while( $err.InnerException ) {
               $err = $err.InnerException
                Write-Host "InnerException: $($err.Message)" 
            }
        }
        else {
            throw $err.Message
        }

    }


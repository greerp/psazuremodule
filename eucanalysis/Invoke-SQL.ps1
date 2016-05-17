<#
  .SYNOPSIS
  Simple SQL Server function to return the results of a query

  .PARAMETER dataSource
  Host name, corresponds to -s on ISQL
  
  .PARAMETER database
  Database Name, corresponds to -d on ISQL

  .PARAMETER sqlCommand
  Query to execute


  .Example 1
  $val = Invoke-Sql -sqlCommand "select Forename + ' ' + Surname as name, dob,address from members"
  foreach ( $row in $val[0].Rows){
    #Convert to array (Not very useful)
    $i = $row.ItemArray

    # Explicitly access row attributes by name
    write-host "Row:" $row.name $row.address $row.dob
  }

  .Example 2
  (Invoke-Sql -sqlCommand "select Forename + ' ' + Surname as name, dob,address from members")[0].Rows|
  %{write-host $_.name }

  #>

function Invoke-SQL {
    param(
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "sdw",
        [string] $sqlCommand = $(throw "Please specify a query."),
        [string] $username = $null,
        [string] $password = $null     
        )

    if ( $username -ne $null ){
        $connectionString = "Data Source=$dataSource;User Id=$username;Password=$password;Initial Catalog=$database"
    }
    else {
        $connectionString = "Data Source=$dataSource;Integrated Security=SSPI;Initial Catalog=$database"
    }

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables

}




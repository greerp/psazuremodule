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

  remove-module Create-SqlObject
  import-module .\Create-SqlObject.ps1
  dir|Create-SqlObject -tableName test

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

function Create-SqlObject {
    param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] $value, 
        [parameter(Mandatory=$true)][String]$tableName=$False,
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "sdw",
        [string] $username,
        [string] $password
        )
    begin {

        ###########################################################################
        function CreateTable($objProps, $name){

            $table = New-Object System.Data.DataTable($name)
            #$members = $obj|Get-Member -MemberType Property
            foreach ( $prop in $objProps ){
                $col = new-object System.Data.DataColumn
                $col.DataType = $prop.TypeName.GetType()
                $col.ColumnName = $prop.Name
                $table.Columns.Add($col)
            }
            <#
                The [ref] otherwise it tries to return a value result and then it gets destroyed
                You have to dereference it using .value 
            #>
            return [ref]$table
        }

        




        ###########################################################################
        if ( $username  ){
            $connectionString = "Data Source=$dataSource;User Id=$username;Password=$password;Initial Catalog=$database"
        }
        else {
            $connectionString = "Data Source=$dataSource;Integrated Security=SSPI;Initial Catalog=$database"
        }
        try {
            $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
            $connection.Open()
        }
        catch {
            throw $_.Exception.Message
        }


        ###########################################################################

        # Variables need to be declared in here if they are retained during processing the stream
        $initialized = $false
        $count = 0 
        $dataTable =$null
        $objProps = $null

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Connecting to ${dataSource}\${database}"
        }

    }

    process {
        $item = $PSitem
        

        if (  -not $initialized ){
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Creating table" $tableName
            }
            $initialized = $true
            $objProps = $item|Get-Member -MemberType Property
            $dataTable = (CreateTable $objProps $tableName).Value
        }

        $count++;
        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Processing Item" $count
        }

        # Add Row

        $dataRow = $dataTable.NewRow()
        if ( $objProps.Length -gt 0 ) {
            foreach ( $prop in $objProps ) {
                $dataRow.Item($prop.Name) = $item.($prop.Name)
                #write-host "Property:" $prop.Name "=" $item.($prop.Name)
            }
            $dataTable.Rows.Add($dataRow)
            [void]$dataRow.AcceptChanges
        }
    }

    end {

        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($connection)
        $bulkCopy.DestinationTableName = $tableName
        $bulkCopy.WriteToServer($dataTable)
        $connection.Close()
    }
}
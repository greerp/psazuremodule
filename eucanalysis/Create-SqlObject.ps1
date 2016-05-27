<#
  .SYNOPSIS
  Given a stream of objects, converts to corresponding SQL columns and persists in existing SWQL table

  .PARMETER value (from Pipeline)
  Object stream of objects to persist in DB

  .PARAMETER tableName
  Existing SQL tablename

  .PARAMETER dataSource
  Host name, corresponds to -s on ISQL
  
  .PARAMETER database
  Database Name, corresponds to -d on ISQL

  .PARAMETER sqlCommand
  Query to execute


  .Example 1

  remove-module Create-SqlObject
  import-module .\Create-SqlObject.ps1

  
  $loc|%{$p= @{id=[int]$_.fileid;loc=[int]$_.loc};  `
       new-object psobject -Property $p}| `
       Create-SqlObject  -tableName "vbaprot" -dataSource "r2gsqlsrv.database.windows.net" -username r2gadmin -password R3dpixie -database redi2go -Verbose

  Note:
  1. You need to pass a proper object stream not a hashtable, if you just pass in @{prpo1=val;prop2=val2} that is a hashtable
  2. SQL table needs to exist. Make sure object property types are the correct ones (note casting in example)


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
                $col.DataType = $prop.Type
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

            
            $objProps = $item|Get-Member -MemberType NoteProperty| `
                              %{$p=$_;New-Object psobject -Property @{Name=$p.Name; Type=$item.($p.Name).GetType()} }


#            $objProps = $item|Get-Member -MemberType NoteProperty 
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

        foreach ($prop in $objProps){
            $columnMap = New-Object `
                System.Data.SqlClient.SqlBulkCopyColumnMapping($prop.Name, $prop.Name)
            $bulkCopy.ColumnMappings.Add($columnMap)
        }

        $bulkCopy.WriteToServer($dataTable)
        $connection.Close()
    }
}

<#
.SYNOPSIS
  Scans folder for Excel file sand outputs VB COM references 

.PARAMETER path
  Foplder to scan
  
.PARAMETER 
  A hash table of key value pairs

#>

Param(
    [String]$file
)


$objExcel = New-Object -ComObject Excel.Application

try {
    $wb = $objExcel.WorkBooks.Open($file, 0, $true,2,"BlahBlahBlah")

    $vba = $wb.VBProject

    foreach ( $ref in $vba.References ){
        $fileRefProps = @{
            OfficeFile = $file
            ComName    = $ref.Name
            ComPath    = $ref.FullPath
            BuiltIn    = $ref.BuiltIn
            IsBroken   = $ref.IsBroken
            Comment    = ""
        }
        
        $officeFile = New-Object psobject -Property $fileRefProps
        Write-Output $officeFile
    }
}
catch {

        Write-Output = New-Object psobject -Property @{
            OfficeFile = $file
            ComName    = ""
            ComPath    = ""
            BuiltIn    = ""
            IsBroken   = ""
            Comment    = $_
        }
}

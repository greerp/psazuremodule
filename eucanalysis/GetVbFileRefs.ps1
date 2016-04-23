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
$wb = $objExcel.WorkBooks.Open($file)
$vba = $wb.VBProject

foreach ( $ref in $vba.References ){
    $fileRefProps = @{
        OfficeFile = $item.FullName
        ComName    = $ref.Name
        ComPath    = $ref.FullPath
        BuiltIn    = $ref.BuiltIn
        IsBroken   = $ref.IsBroken
    }
        
    $officeFile = New-Object psobject -Property $fileRefProps
    Write-Output $officeFile
}

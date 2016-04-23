<#
.SYNOPSIS
  Scans folder for Excel file sand outputs VB COM references 

.PARAMETER path
  Foplder to scan
  
.PARAMETER 
  A hash table of key value pairs

#>

Param(
    [String]$path = "f:\temp"
)

$output = @()
$ext = @("*.xls","*.xlt")

$objExcel = New-Object -ComObject Excel.Application

ForEach ($item in (Get-ChildItem -Path $path -Recurse -Include $ext)) {

    #Write-Output $item.FullName

    $wb = $objExcel.WorkBooks.Open($item.FullName)
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
    
#    $wb.VBProject.VBComponents
#    foreach ( $codeComp in $wb.VBProject.VBComponents){
#    }
}

#$output |
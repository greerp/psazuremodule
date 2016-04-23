<#
  .SYNOPSIS
  Iterates a csv file with .xls/.xlsx and outputs the VBA References 

  .PARAMETER csvfile
  Input containing excel file names
  
  .PARAMETER outfile
  Output CSV file 

  #>

  Param (
    [String]$csvfile,
    [String]$outfile
  )


  function ifEmpty($toCheck, $value){
     if ([string]::IsNullOrWhiteSpace($toCheck)){
        return $value
     }
     else {
        return $toCheck
     }


  }

  function getExcelInstance {
     $objExcel = New-Object -ComObject Excel.Application
     $objExcel.DisplayAlerts = $False
     $objExcel.EnableEvents  = $False
     $objExcel.Interactive   = $False
     $objExcel.AutomationSecurity = 
            [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable

     return $objExcel
  }


  #Configure Excel to be headless

  $objExcel = getExcelInstance

  #$objExcel.Visible = $True

  # Read input files names into a collection
  $files = import-csv $csvfile -Header fileid,filepath,skip
  $item = 1
  $itemCount = $files.Count


  foreach ($file in $files){
    # Skip non excel and files and files marked in csv attribute "skip"
    if ( $file.skip -eq "" -and 
         (($file.filepath.ToLower().EndsWith(".xls")) -or ($file.filepath.ToLower().EndsWith(".xlsm"))) ) { 

        Write-Host "Processing $item of $itemCount files - id: $($file.fileid)"

        $wb=$null
        try {
            # Specify password so that if the file has one, it causes an exception rather than prompt
            $wb = $objExcel.WorkBooks.Open($file.filepath, 0, $true, 2, "BlahBlaBlah")

            if ($wb.HasVBProject) {
                $vba = $wb.VBProject

                foreach ( $ref in $vba.References ){
                    $fileRefProps = @{
                        FileId     = $file.fileid
                        ComName    = $ref.Name
                        Guid       = $ref.Guid
                        ComPath    = $ref.FullPath
                        BuiltIn    = $ref.BuiltIn
                        IsBroken   = $ref.IsBroken
                        Comment    = ""
                    }
                    $officeFile = New-Object psobject -Property $fileRefProps
                    $officeFile | export-csv -Path $outfile -append -NoTypeInformation
                }
            }
        }
        catch {
            $msg = $_ -replace '[\W]',''
            New-Object psobject -Property @{
                FileId     = $file.fileid
                ComName    = ""
                Guid       = ""
                ComPath    = ""
                BuiltIn    = ""
                IsBroken   = ""
                Comment    = $msg } | 
            export-csv -Path $outfile -append -NoTypeInformation -force

        }
        finally {
            # Excel instance killed - Recreate
            if ($objExcel.Application -eq $null ){
                $objExcel = getExcelInstance
            }

            elseif ($wb) { 
                $wb.Close($false)
            }
            $item++
        }
    }



}



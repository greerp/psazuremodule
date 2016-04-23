  <#
  .SYNOPSIS
  Iterates a csv file with .xls/.xlsx and outputs the VBA References 

  .PARAMETER csvfile
  Input containing excel file names
  
  .PARAMETER outfile
  Output CSV file 

  #>

  Param(
    [String]$csvfile,
    [String]$outfile
  )

  #Configure Excel to be headless
  $objExcel = New-Object -ComObject Excel.Application
  $objExcel.DisplayAlerts = $False
  $objExcel.EnableEvents  = $False
  $objExcel.Interactive   = $False
  $objExcel.AutomationSecurity = 
    [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable

  #$objExcel.Visible = $True

  # Read input files names into a collection
  $files = import-csv $csvfile -Header fileid,filepath
  $item = 1
  $itemCount = $files.Count


  foreach ($file in $files){
    if ( ($file.filepath.ToLower().EndsWith(".xls")) -or ($file.filepath.ToLower().EndsWith(".xlsx")) ) { 
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
                        FileName   = $file.filepath
                        ComName    = $ref.Name
                        ComPath    = $ref.FullPath
                        BuiltIn    = $ref.BuiltIn
                        IsBroken   = $ref.IsBroken
                        Comment    = ""
                    }
                    $officeFile = New-Object psobject -Property $fileRefProps
                    $officeFile | export-csv -Path $outfile -append -NoTypeInformation
                }
            }
            $item++
        }
        catch {
            New-Object psobject -Property @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                ComName    = ""
                ComPath    = ""
                BuiltIn    = ""
                IsBroken   = ""
                Comment    = $_ } | 
            export-csv -Path $outfile -append -NoTypeInformation -force

        }
        finally {
            if ($wb) { 
                $wb.Close($false)
            }

        }
    }
}

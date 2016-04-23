<#
  .SYNOPSIS
  Iterates a csv file with .xls/.xlsx and couns lines of VBA Code 

  .PARAMETER csvfile
  Input containing excel file names
  
  .PARAMETER outfile
  Output CSV file 

  #>

    Param (
        [String]$csvfile,
        [String]$outfile,
        [parameter(Mandatory=$false)][String[]]$contains,
        [parameter(Mandatory=$false)][bool]$ScanMDB=$False
    )

    # System Values
    
    # Excel extentsions to process
    $xlsExt = '.xls' '.xlsb' '.xlsx' '.xlxm' 
    #Access extensions to process
    $mdbExt = '.mdb' '.mde'



    ###########################################################################
    function getExcelInstance {
        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.DisplayAlerts = $False
        $objExcel.EnableEvents  = $False
        $objExcel.Interactive   = $False
        $objExcel.AutomationSecurity = 
            [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable
        return $objExcel
    }

    ###########################################################################
    function getAccessInstance {
        $objAccess = New-Object -ComObject Access.Application
        $objAccess.DisplayAlerts = $False
        $objAccess.EnableEvents  = $False
        $objAccess.Interactive   = $False
        $objAccess.AutomationSecurity = 
            [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable
        return $objAccess
    }

    ###########################################################################
    function ProcessExcelFile($file, $item, $itemCount, $outfile, $objExcel) {
        $wb=$null
        try {
            if ( -not (test-Path $file.filepath) ){
                throw [System.IO.FileNotFoundException] "File Missing"
            }

            # Specify password so that if the file has one, it causes an exception rather than prompt
            $wb = $objExcel.WorkBooks.Open($file.filepath, 0, $true, 2, "BlahBlaBlah")

            $totLoc = 0 
            $containsToken=$False
            if ($wb.HasVBProject) {
                $vba = $wb.VBProject

                foreach ( $comp in $vba.VBComponents ){
                    $loc = $comp.CodeModule.CountOfLines
                    $totLoc+=$loc

                    if ($loc -gt 0 -and $containsToken -eq $False -and $Contains.Count -gt 0){
                
                        $code = $comp.CodeModule.Lines(1, $loc )
                        foreach ( $token in $contains ){
                            if ( $code -match $token ){
                                $containsToken = $true
                                break
                            }
                        }                    
                    }
                }
            }

            $fileRefProps = @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                Loc        = $totLoc
                Contains   = $containsToken
                Comment    = ""
            }
            $officeFile = New-Object psobject -Property $fileRefProps
            $officeFile | export-csv -Path $outfile -append -NoTypeInformation
            
            if ($PSBoundParameters['Verbose']){
                Write-Host "Processed $item of $itemCount files - id: $($file.fileid), LOC=$totloc, Contains Token=$containsToken"
            }
        }
        catch {
        #Todo - Test for password specific exception 

            $msg = $_
                New-Object psobject -Property @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                Loc        = $null
                Contains   = $false
                Comment    = $msg } | 
            export-csv -Path $outfile -append -NoTypeInformation -force
            Write-Host "Error Processing $item of $itemCount files - id: $($file.fileid), see report..."
        }
        finally {
            if ($wb) { 
                $wb.Close($false)
            }
        }
    }

    ###########################################################################
    function ProcessAccessFile($file, $item, $itemCount, $outfile, $objAccess) {
        $db=$null
        try {
            if ( -not (test-Path $file.filepath) ){
                throw [System.IO.FileNotFoundException] "File Missing"
            }

            # Specify password so that if the file has one, it causes an exception rather than prompt
            $db = $objAccess.OpenCurrentDatabase($file.filepath, $False, "BlahBlaBlah")
            $totLoc = 0 
            $containsToken=$False

            foreach ($vbp in $db.VBE.VBProjects){
                foreach ( $vbc in $vbp.VBComponents){
                    $loc = $vbc.CodeModule.CountOfLines
                    $totLoc+=$loc

                    if ($loc -gt 0 -and $containsToken -eq $False -and $Contains.Count -gt 0){
                        $code = $vbc.CodeModule.Lines(1, $loc )
                        foreach ( $token in $contains ) {
                            if ( $code -match $token ) {
                                $containsToken = $true
                                break
                            }
                        } 
                    }
                }
            }

            $fileRefProps = @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                Loc        = $totLoc
                Contains   = $containsToken
                Comment    = ""
            }
            $officeFile = New-Object psobject -Property $fileRefProps
            $officeFile | export-csv -Path $outfile -append -NoTypeInformation
            
            Write-Host "Processed $item of $itemCount files - id: $($file.fileid), LOC=$totloc, Contains Token=$containsToken"
        }
        catch {
            $msg = $_
                New-Object psobject -Property @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                Loc        = $null
                Contains   = $false
                Comment    = $msg } | 
            export-csv -Path $outfile -append -NoTypeInformation -force
            Write-Host "Error Processing $item of $itemCount files - id: $($file.fileid), see report..."
        }
        finally {
            if ($db) { 
                $db.CloseCurrentDatabase
            }
        }
    }


    ###########################################################################
    $objExcel = getExcelInstance
    if ( $ScanMDB ) {
        $objAccess = getAccessInstance
    }

    $files = import-csv $csvfile -Header fileid,filepath,skip -Encoding UTF7
    $item = 1
    $itemCount = $files.Count

    foreach ($file in $files){
        if ( [string]::IsNullOrEmpty($file.skip) ) {
          $fileExt =  [System.IO.Path]::GetExtension($file.filepath)
          if ( $xlsExt -match $fileExt ) { 
              ProcessExcelFile $file, $item $itemCount $outFile $objExcel

              # Excel instance killed - Recreate
              if ($objExcel.Application -eq $null ){
                  $objExcel = getExcelInstance
              }
          }
          elseif ( $ScanMDB -and $xlsExt -match $fileExt ) { 
              ProcessAccessFile $file $item $itemCount $outFile $objAccess

              # Excel instance killed - Recreate
              if ($objExcel.Application -eq $null ){
                  $objExcel = getExcelInstance
              }
          }
        }
        $item++
    }



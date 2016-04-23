<#
  .SYNOPSIS
  Accepts pipeline of Excel and Access files and outputs ActiveX References

  .PARAMETER input
  Pipeline objects containing id and filepath properties 
  

  .Example
  import-module .\GetActiveXReferences.ps1
  dir *.xls|%{$i=0}{@{id=$i++;filepath=$_.FullName}}|Get-ComRefs
  remove-module GetActiveXReferences


  #>

function Get-ComRefs {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] $value, 
        [parameter(Mandatory=$false)][bool]$ScanMDB=$False

    )
  
    begin {
        ###########################################################################
        # Inner Functions
        function ifEmpty($toCheck, $value){
            if ([string]::IsNullOrWhiteSpace($toCheck)){
            return $value
            }
            else {
            return $toCheck
            }
        }

        ###########################################################################
        function getExcelInstance {
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Creating instance of Excel for automation"
            }
            $objExcel = New-Object -ComObject Excel.Application
            $objExcel.DisplayAlerts = $False
            $objExcel.EnableEvents  = $False
            $objExcel.Interactive   = $False
            #$objExcel.AutomationSecurity = 
            #    [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable
            return $objExcel
        }

        ###########################################################################
        function getAccessInstance {
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Creating instance of Access for automation"
            }
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

                # Initialisers for output object 
                $totLoc = 0 
                $containsToken=$False
                $refs=@()

                if ($wb.HasVBProject) {
                    $vba = $wb.VBProject

                    #Scan References
                    foreach ( $vbaRef in $vba.References ){
                        $fileRefProps = @{
                            ComName    = $vbaRef.Name
                            Guid       = $vbaRef.Guid
                            ComPath    = $vbaRef.FullPath
                            BuiltIn    = $vbaRef.BuiltIn
                            IsBroken   = $vbaRef.IsBroken
                            Tight      = $True
                        }
                        # Add each ref to an array
                        $refs.Add((New-Object psobject -Property $fileRefProps))
                    }
                    
                    # Scan VBA Code
                    foreach ( $comp in $vba.VBComponents ) {
                        $loc = $comp.CodeModule.CountOfLines
                        $totLoc+=$loc

                        if ($loc -gt 0 -and $containsToken -eq $False -and $Contains.Count -gt 0) {
                            # Look for token
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
                    References = $refs
                    Comment    = ""
                }
                
                # Output value
                if ($PSBoundParameters['Verbose']){
                    Write-Host "Processed $item of $itemCount files - id: $($file.fileid), LOC=$totloc, References Count:$($refs.Count)"
                }

                New-Object psobject -Property $fileRefProps
            }
            catch {
                if ($PSBoundParameters['Verbose']){
                    Write-Host "Error Processing $item of $itemCount files - id: $($file.fileid), see report..."
                }

                #Todo - Test for password specific exception 
                $msg = $_ -replace '[\W]',''
                New-Object psobject -Property @{
                FileId     = $file.fileid
                FileName   = $file.filepath
                Loc        = $null
                Contains   = $false
                References = $null
                Comment    = $msg } 
            }
            finally {
                if ($wb) { 
                    $wb.Close($false)
                }
            }
        }

        ##### REAL BEGINNING #####

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Initialising...."
        }
        $objExcel = getExcelInstance
        if ( $ScanMDB ) {
            $objAccess = GetAccessInstance
        }


    }

    ###########################################################################
    Process {
        $file = @{
            fileid=$value.id;
            filepath=$value.filepath}

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host  "Processing " $file.filepath
        }

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
    }
  
    ###########################################################################
    End {
        Write-host "In End Block"
        $objExcel.Close
        $objAccess.Close
    }
}



<#
  .SYNOPSIS
  Accepts pipeline of Excel and Access files and outputs ActiveX References

  .PARAMETER input
  Pipeline objects containing id and filepath properties 
  

  .Example 1
  import-module .\Get-ComRefs.ps1
  dir *.xls|%{$i=0}{@{id=$i++;filepath=$_.FullName}}|Get-ComRefs
  remove-module Get-Comrefs

  .Example 2
  @{id=1;filepath="C:\Users\greepau\Desktop\Order Log Monthly_Template.xls"}|Get-ComRefs

  .Example 3
  $d = dir *.xls
  $d|%{$i=0}{@{id=$i++;filepath=$_.FullName}}|
    Get-ComRefs|
    %{$f=$_.filename;foreach ($r in $_.References){Write-Host $f $r.ProgId}}

  .Example 4
  $d|%{$i=0}{@{id=$i++;filepath=$_.FullName}}|
    Get-ComRefs|
    %{write-host "File:" $_.filename ", Loc:" $_.Loc ", Hash:" $_.Codehash}



  .Notes
  ERRORS IMPORTANT

  No protected view Window 
  ========================
  1-Go into Excel - File/Options/Trust Center/Trust Center settings/Protected View
  2-Disable the protected view settings 
  
  Old Format or Invalid Type Library
  ==================================
  On the test machine ensure that within the region settings 
  BOTH Formats and Location are set to United Kingdom  


  Changes
  =======
  05/05/16 - Added vba protected test

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

<#
        function UnlockVbaProject($objExcel, $wbToUnlock) {

            # Load the unlock macro code from a separate .ps1
            $code = . .\UnlockVbaMacro.ps1

            $vbaWb = $objExcel.WorkBooks.Add()
            $module = $vbaWb.VBProject.VBComponents.Add(1)
            $module.CodeModule.AddFromString($code)

            #Unlock VBA
            $wbToUnlock.Activate
            $macro= $vbaWb.Name + "!Unprotect"
            $objExcel.Application.Run($macro)

            $vbaWb.Close($false)


            return
        }
#>
        ###########################################################################
        function getExcelInstance {
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Creating instance of Excel for automation"
            }
            $objExcel = New-Object -ComObject Excel.Application
            #$objExcel.DisplayAlerts = $False
            #$objExcel.EnableEvents  = $False
            $objExcel.Interactive   = $True
            try {
                $objExcel.AutomationSecurity = 
                    [Microsoft.Office.Core.MsoAutomationSecurity]::msoAutomationSecurityForceDisable 
            }
            catch {}

            return $objExcel
        }

        ###########################################################################
        function getAccessInstance() {
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
        Function Get-StringHash([String] $String,$HashName = "MD5")
        {
            $StringBuilder = New-Object System.Text.StringBuilder
            [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))|%{
                [Void]$StringBuilder.Append($_.ToString("x2"))
            }
            $StringBuilder.ToString()
        }

        ###########################################################################
        function getLateBoundObjects($source) {

            $lateBoundObjects =[System.Collections.arrayList]@()
            #$lateBoundObjects =@()
            # regex Expr, note ?<progid> which names the group
            #$regexComRef = "CreateObject\([`"'](?<progid>[\w\.]+)[`"']\)"
            $regexComRef = "(?:Create|Get)Object\([`"'](?<progid>[\w\.]+)[`"']\)"

            # Use Select String to do multiple macthes in a string, pick out the group value which is denoted in the regex by \w+.\w+
            $progids = select-string -inputObject $source -Pattern $regexComRef -AllMatches | 
                % { $_.Matches } | 
                % { $_.Groups['progid'].Value }

            foreach ( $progid in $progids ){
                $type = [System.Type]::GetTypeFromProgID($progid, $false)

                $fileRefProps = @{
                    ComName    = $null
                    ProgId     = $progid
                    Guid       = $null
                    ComPath    = $null
                    BuiltIn    = $false
                    IsBroken   = $true
                    LateBound  = $True
                }

                if ( $type -ne $null ) {
                    $fileRefProps.ComName = $type.FullName
                    $filerefProps.Guid = $type.Guid
                    $fileRefProps.IsBroken = $false
                }
                #$obj = New-Object psobject -Property $fileRefProps
                #$lateBoundObjects+=$obj


                [void]$lateBoundObjects.Add(
                    (New-Object psobject -Property $fileRefProps))
                
            }
            return $lateBoundObjects
        }

        ###########################################################################
        function ProcessExcelFile($file, $objExcel) {
            $wb=$null
            try {
                if ( -not (test-Path $file.filepath) ){
                    throw [System.IO.FileNotFoundException] "File Missing"
                }

                # Specify password so that if the file has one, it causes an exception rather than prompt
                $wb = $objExcel.WorkBooks.Open($file.filepath, 0, $true, 2, "BlahBlaBlah")

                # Initialisers for output object 
                $containsToken=$False
                $refs=[System.Collections.arrayList]@()
                $hash=""
                $vbaProt = $null

                if ($wb.HasVBProject) {
                    $vba = $wb.VBProject
                    $vbaProt=$vba.Protection


                    ###########################################
                    #Scan Tightly Bound References
                    foreach ( $vbaRef in $vba.References ){
                        $fileRefProps = @{
                            ComName    = $vbaRef.Name
                            ProgId     = $null
                            Guid       = $vbaRef.Guid
                            ComPath    = $vbaRef.FullPath
                            BuiltIn    = $vbaRef.BuiltIn
                            IsBroken   = $vbaRef.IsBroken
                            LateBound  = $false
                        }
                        [void]$refs.Add((New-Object psobject -Property $fileRefProps))
                    }
                    
                    ###########################################
                    # Scan VBA Code
                    $moduleHashes=""
                    if (  $vbaProt -eq 0 ){
                        $totLoc = 0
                        foreach ( $comp in $vba.VBComponents ) {
                            $loc = $comp.CodeModule.CountOfLines
                            $totLoc+=$loc

                            if ( $loc -gt 0 ) {
                                $code = $comp.CodeModule.Lines(1, $loc )
                                $moduleHashes += (Get-StringHash -String $code)

                                ###########################################
                                # Scan Code for COM Refs
                                $lateRefList = @(getLateBoundObjects $code)
                                if ( $lateRefList.Count -gt 0){
                                    foreach($lateRef in $lateRefList){
                                        [void]$refs.Add($lateRef)
                                    }
                                }

                                if ($containsToken -eq $False -and $Contains.Count -gt 0) {
                                    # Look for token
                                    foreach ( $token in $contains ){
                                        if ( $code -match $token ){
                                            $containsToken = $true
                                            break
                                        }
                                    }                    
                                }
                            }
                        }
                    }
                    if ( $moduleHashes -ne "" ) {
                        $hash = (Get-StringHash -String $moduleHashes)
                    }

                }

                $fileRefProps = @{
                    FileId     = $file.fileid
                    FileName   = $file.filepath
                    Loc        = $totLoc
                    CodeHash   = $hash
                    Contains   = $containsToken
                    References = $refs
                    VbaProt    = $vbaProt
                    Comment    = ""
                }
                

                New-Object psobject -Property $fileRefProps
            }
            catch {
                #Todo - Test for password specific exception 
                $msg = $_ -replace '[\n]',' '
                New-Object psobject -Property @{
                    FileId     = $file.fileid
                    FileName   = $file.filepath
                    Loc        = $null
                    CodeHash  = $null
                    Contains   = $false
                    References = $null
                    VbaProt    = $null
                    Comment    = $msg } 
            }
            finally {
                if ($wb) { 
                    $wb.Close($false)
                }
            }
        }

        ##### REAL BEGINNING #####
        # Excel extentsions to process
        $xlsExt = ('.xls','.xlsb','.xlsx','.xlsm','.xlt')
        #Access extensions to process
        $mdbExt = ('.mdb','.mde')


        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Initialising...."
        }
        $objExcel = getExcelInstance
        $count = 0
        if ( $ScanMDB ) {
            $objAccess = GetAccessInstance
        }


    }

    ###########################################################################
    Process {
        $file = @{
            fileid=$value.id;
            filepath=$value.filepath}

        [Reflection.Assembly]::LoadFile("${PSScriptRoot}\euclib.dll")
        $vbaProps = [com.redpixie.euc.mcdf.PsWrapper]::GetDocumentProperties($file.filepath)


        $count++;

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host  "Processing EUC file $count, " $file.filepath
        }

        if ( [string]::IsNullOrEmpty($file.skip) ) {
            $fileExt =  [System.IO.Path]::GetExtension($file.filepath)
            if ( $xlsExt -match $fileExt ) { 
                ProcessExcelFile $file $objExcel

                # Excel instance killed - Recreate
                if ($objExcel.Application -eq $null ){
                    $objExcel = getExcelInstance
                }
            }
            elseif ( $ScanMDB -and $xlsExt -match $fileExt ) { 
                ProcessAccessFile $file $objAccess

                # Excel instance killed - Recreate
                if ($objExcel.Application -eq $null ){
                    $objExcel = getExcelInstance
                }
            }
        }
        else {

        }
    }
  
    ###########################################################################
    End {
        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host "Terminating"
        }
        $objExcel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objExcel)
        if ( $ScanMDB -and $objAccess.Application -eq $null) {
            $objAccess.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objAccess)
        }
    }
}



<#
  .SYNOPSIS
  Accepts pipeline of Excel and Access files and outputs ActiveX References

  .PARAMETER input
  Pipeline objects containing id and filepath properties 
  

  .Example 1
  import-module .\Get-OfficeDocProps.ps1
  dir *.xls|%{$i=0}{@{id=$i++;filepath=$_.FullName}}|Get-OfficeDocProps
  remove-module Get-OfficeDocProps

  .Example 2
   @{id=1;filepath="C:\Users\greepau\Desktop\test.xls"}|Get-OfficeDocProps
  
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

function Get-OfficeDocProps {
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
        function LoadMcdfLibrary {
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Loading MSOVBA Library"
            }
            [Reflection.Assembly]::LoadFile("${PSScriptRoot}\euclib.dll")

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
        function getComAttributes($clsid){
            
            $foundCls = $false
            $regPaths = @("HKLM\SOFTWARE\Classes\Wow6432Node\CLSID\","HKLM:\SOFTWARE\CLASSES\CLSID\")
            $name = $null
            $server = $null

            foreach ( $clsRoot in $regPaths){
                $tryPath = "${clsRoot}${clsId}"
                if (($foundCls -eq $false) -and (Test-Path $tryPath) ) {
                    $name = (Get-ItemProperty $tryPath)."(default)"

                    $tryPath = "${clsRoot}${clsId}\InProcServer32"
                    if (Test-Path $tryPath){
                        $server = (Get-ItemProperty $tryPath)."(default)"
                    }
                    else {
                        $tryPath = "${clsRoot}${clsId}\LocalServer32"
                        if (Test-Path $tryPath){
                            $server = (Get-ItemProperty $tryPath)."(default)"
                        }
                    }
                    $foundCls = $true                    
                }
            }
            return @{name=$name;inprocserver32=$inprocserver32}
        }

        ###########################################################################
        function ProcessExcelFile($file) {
            $wb=$null
            try {
                if ( -not (test-Path $file.filepath) ){
                    throw [System.IO.FileNotFoundException] "File Missing"
                }

                $vbaProps = [com.redpixie.euc.mcdf.PsWrapper]::GetDocumentProperties($file.filepath)

                # Initialisers for output object 
                $refs=[System.Collections.arrayList]@()

                if ($vbaProps.TotLoc -gt 0 ) {
                    $hash = $vbaProps.ModuleHash
                    $totLoc = $vbaProps.TotLoc

                    ###########################################
                    #Scan Tightly Bound References
                    foreach ( $vbaRef in $vbaProps.EarlyBoundReferences ){
                        $fileRefProps = @{
                            ComName    = $vbaRef.Name
                            ProgId     = $null
                            Guid       = $vbaRef.Guid
                            ComPath    = $vbaRef.Path
                            BuiltIn    = $null
                            IsBroken   = $null
                            LateBound  = $false
                            RefCount   = $null
                        }
                        [void]$refs.Add((New-Object psobject -Property $fileRefProps))
                    }

                    ###########################################
                    # Check for contains tokens
                    if ($containsToken -eq $False -and $Contains.Count -gt 0) {
                        # Look for token
                        foreach ( $token in $contains ){
                            foreach ($module in $vbaProps.Modules) {
                                if ( $module.Loc -gt 0 ) {
                                    $code = $module.Code
                                    if ( $code -match $token ){
                                        $containsToken = $true
                                        break
                                    }
                                }
                            }
                        }                    
                    }

                    ###########################################
                    # Add the late bound progids

                    foreach ($ref in $vbaProps.LateBoundReferences.GetEnumerator()) {

                        $ComName = $null
                        $Guid = $null
                        $IsBroken = $true

                        # Get COM info for ProgID
                        $type = [System.Type]::GetTypeFromProgID($ref.Key, $false)

                        if ( $type -ne $null ) {
                            $comAttr = getComAttributes $type.Guid
                            $comName = $comAttr.Name
                            $comPath = $comAttr.Inprocserver32
                            $Guid = $type.Guid
                            $IsBroken = $false
                        }

                        $fileRefProps = @{
                            ComName    = $ComName
                            ProgId     = $ref.Key
                            Guid       = $Guid
                            ComPath    = $ComPath
                            BuiltIn    = $false
                            IsBroken   = $IsBroken
                            LateBound  = $true
                            RefCount   = $ref.Value
                        }
                        [void]$refs.Add((New-Object psobject -Property $fileRefProps))
                    }

                }

                $fileProps = @{
                    FileId     = $file.fileid
                    FileName   = $file.filepath
                    Loc        = $totLoc
                    CodeHash   = $hash
                    Contains   = $containsToken
                    References = $refs
                    VbaProt    = $null
                    Comment    = ""
                    Modules    = $vbaProps.Modules
                }
                

                New-Object psobject -Property $fileProps
            }
            catch {
                #Todo - Test for password specific exception 
                $msg = $_ -replace '[\n]',' '
                New-Object psobject -Property @{
                    FileId     = $file.fileid
                    FileName   = $file.filepath
                    Loc        = $null
                    CodeHash   = $null
                    Contains   = $false
                    References = $null
                    VbaProt    = $null
                    Comment    = $msg 
                    Modules    = $null  
               } 
            }
            finally {
                if ($wb) { 
                    $wb.Close($false)
                }
            }
        }

        ##### REAL BEGINNING #####
        # Excel extentsions to process
        $oleExt = ('.xls','.xlt','.doc','.xlsm')
        #Access extensions to process
        $mdbExt = ('.mdb')


        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Initialising...."
        }

        LoadMcdfLibrary
        
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

        $count++;

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host  "Processing EUC file $count, " $file.filepath
        }

        if ( [string]::IsNullOrEmpty($file.skip) ) {
            $fileExt =  [System.IO.Path]::GetExtension($file.filepath)
            if ( $oleExt -match $fileExt ) { 
                ProcessExcelFile $file

            }
            elseif ( $ScanMDB -and $mdbExt -match $fileExt ) { 
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
        if ( $ScanMDB -and $objAccess.Application -eq $null) {
            $objAccess.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objAccess)
        }
    }
}



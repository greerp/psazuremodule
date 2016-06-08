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
  @{id=1;filepath="C:\Users\greepau\Desktop\test.xls"}|Get-ComRefs

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
            $objExcel.DisplayAlerts = $False
            $objExcel.EnableEvents  = $False
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
        function ExcelChartCount($wb) {
            $chartCount = 0

            foreach ( $ws in $wb.worksheets ){
                $charts = $ws.ChartObjects()
                if ( $charts -ne $null  ){
                    $chartCount+=$charts.Count
                }    
            }
            return $chartCount
        }
        ###########################################################################
        function ExcelLinkCount($wb) {
            $linkCount = 0

            foreach ( $link in $wb.LinkSources(1) ){
                $linkCount++
            }
            return $linkCount
        }


        ###########################################################################
        function getLateBoundObjects($source) {

            $lateBoundObjects =[System.Collections.hashtable]@{}
            # regex Expr, note ?<progid> which names the group
            $regexComRef = "(?:Create|Get)Object\([`"'](?<progid>[\w\.]+)[`"']\)"

            # Use Select String to do multiple macthes in a string, pick out the group
            # value which is denoted in the regex by \w+.\w+
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
                    $comAttr = getComAttributes $type.Guid
                    $comName = $comAttr.Name
                    $comPath = $comAttr.Inprocserver32

                    $fileRefProps.ComName = $type.FullName
                    $filerefProps.Guid = $type.Guid
                    $fileRefProps.IsBroken = $false
                }
                $comRef = New-Object psobject -Property $fileRefProps

                if ( $lateBoundObjects.Contains($progid) -ne $false) {
                    [void]$lateBoundObjects.Add($progid,$comRef)
                }
                
            }
            return $lateBoundObjects.Values
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
        function getDocumentProperties($wb){

            $binding = “System.Reflection.BindingFlags” -as [type]
            $objHash = @{}

            Foreach($property in $wb.BuiltInDocumentProperties) {
                try {
                    $propName = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$property,$null)
                    $propValue = [System.__ComObject].invokemember("value",$binding::GetProperty,$null,$property,$null)
                    [void]$objHash.Add($propName,$propValue)
                }
                catch { }
            }
            return $objHash
        }

        ###########################################################################
        function ProcessAccessFile($file, $objAccess) {
            $fileObj = Get-Item -Path $file.filepath
            $fileProps.Add("Owner", (Get-Acl -Path $file.filepath).Owner )
            $fileProps.Add("Size", $fileObj.Length)
            $fileProps.Add("Created", $fileObj.CreationTime)
            $fileProps.Add("Updated", $fileObj.LastWriteTime)

            $fileOwner=$null
            $fileProps=@{}


            $fileRefProps = @{
                FileId         = $file.fileid
                FileName       = split-path $file.filepath -Leaf
                Directory      = split-path $file.filepath
                Loc            = $null
                CodeHash       = $null
                Contains       = $null
                References     = $refs
                VbaProt        = $null
                ChartCount     = $null
                AXCount        = $null
                LinkCount      = $null
                Owner          = $fileProps["Owner"]
                Size           = $fileProps["Size"]
                Created        = $fileProps["Created"]
                Updated        = $fileProps["Updated"]
                Comment    = ""
            }
            New-Object psobject -Property $fileRefProps
        }


        ###########################################################################
        function ProcessExcelFile($file, $objExcel) {
            $wb=$null
            $chartCount = 0
            $AXCount = 0
            $linkCount = 0 
            $docProps = @{}
            $fileOwner=$null
            $fileProps=@{}
            $filePwd = $false
            try {
                if ( -not (test-Path $file.filepath) ){
                    throw [System.IO.FileNotFoundException] "File Missing"
                }

                $fileObj = Get-Item -Path $file.filepath
                $fileProps.Add("Owner", (Get-Acl -Path $file.filepath).Owner )
                $fileProps.Add("Size", $fileObj.Length)
                $fileProps.Add("Created", $fileObj.CreationTime)
                $fileProps.Add("Updated", $fileObj.LastWriteTime)
                

                # Specify password so that if the file has one, it causes an exception rather than prompt
                try {
                    $wb = $objExcel.WorkBooks.Open($file.filepath, 0, $true, 2, "BlahBlaBlah")
                }
                catch {
                    $ex = $_
                    if ( $ex.Exception.Errorcode -eq -2146827284 ){
                        $filePwd = $true
                    }
                    else {
                        throw $ex
                    }
                }


                # Initialisers for output object 
                $containsToken=$False
                $refs=[System.Collections.arrayList]@()
                $hash=""
                $vbaProt = $null

                $chartCount = ExcelChartCount $wb
                $linkCount = ExcelLinkCount $wb

                if ($wb.HasVBProject) {
                    $vba = $wb.VBProject
                    $vbaProt=$vba.Protection

                    ###########################################
                    #Scan Tightly Bound References
                    foreach ( $vbaRef in $vba.References ){
                        if ( $vbaRef.Guid ) {
                            $progid  = (GetProgId $vbaRef.Guid).progid
                        }
                        else {
                            $progid=$null
                        }
                        $fileRefProps = @{
                            ComName    = $vbaRef.Name
                            ProgId     = $progid
                            Guid       = $vbaRef.Guid
                            ComPath    = $vbaRef.FullPath
                            BuiltIn    = $vbaRef.BuiltIn
                            IsBroken   = $vbaRef.IsBroken
                            LateBound  = $false
                        }
                        if ( $vbaRef.BuiltIn -eq $false){
                            $AXCount++
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
                                    $AXCount+=$lateRefList.Count
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
                    FileId         = $file.fileid
                    FileName       = split-path $file.filepath -Leaf
                    Directory      = split-path $file.filepath
                    Loc            = $totLoc
                    CodeHash       = $hash
                    Contains       = $containsToken
                    References     = $refs
                    VbaProt        = $vbaProt
                    ChartCount     = $chartCount
                    AXCount        = $AXCount
                    LinkCount      = $LinkCount
                    Owner          = $fileProps["Owner"]
                    Size           = $fileProps["Size"]
                    Created        = $fileProps["Created"]
                    Updated        = $fileProps["Updated"]
                    FilePwd        = $filePwd
                    Comment    = ""
                }
                New-Object psobject -Property $fileRefProps
            }
            catch {
                #Todo - Test for password specific exception 
                $msg = $_ -replace '[\n]',' '
                New-Object psobject -Property @{
                    FileId     = $file.fileid
                    FileName   = split-path $file.filepath -Leaf
                    Directory  = split-path $file.filepath
                    Loc        = $null
                    CodeHash   = $null
                    Contains   = $false
                    References = [System.Collections.arrayList]@()
                    VbaProt    = $null
                    ChartCount = $null
                    AXCount    = $null
                    LinkCount  = $null
                    Owner      = $fileProps["Owner"]
                    Size       = $fileProps["Size"]
                    Created    = $fileProps["Created"]
                    Updated    = $fileProps["Updated"]
                    FilePwd    = $filePwd
                    Comment    = $msg 
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
        $xlsExt = ('.xls','.xlsb','.xlsx','.xlsm','.xlt')
        #Access extensions to process
        $mdbExt = ('.mdb')


        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Initialising...."
        }
        # Load GetProgId function 
        . .\GetProgid.ps1
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

        $count++;

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host  "Processing EUC file $count, " $file.filepath
        }

        if ( [string]::IsNullOrEmpty($file.skip) ) {
            $fileExt =  [System.IO.Path]::GetExtension($file.filepath)
            if ( $fileExt -eq "" ){
                # Skip
            }
            elseif ( $xlsExt -match $fileExt ) { 
                ProcessExcelFile $file $objExcel

                # Excel instance killed - Recreate
                if ($objExcel.Application -eq $null ){
                    $objExcel = getExcelInstance
                }
            }
            elseif ( $ScanMDB -and $xlsExt -match $fileExt ) { 
                ProcessAccessFile $file $objAccess

                # Excel instance killed - Recreate
                if ($objAccess.Application -eq $null ){
                    $objAccess = getAccessInstance
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
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($objExcel)
        if ( $ScanMDB -and $objAccess.Application -eq $null) {
            $objAccess.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($objAccess)
        }
    }
}



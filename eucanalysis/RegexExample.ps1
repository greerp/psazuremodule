# Match ProgIds within a string 

        Function Get-StringHash([String] $String,$HashName = "MD5")
        {
            $StringBuilder = New-Object System.Text.StringBuilder
            [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash([System.Text.Encoding]::UTF8.GetBytes($String))|%{
                [Void]$StringBuilder.Append($_.ToString("x2"))
            }
            $StringBuilder.ToString()
        }

# Source code 
$source = "set x = CreateObject('Excel.Application') set y=CreateObject(`"DFSRHelper.ServerHealthReport.15`")"

# regex Expr, note ?<progid> which names the group
$regexComRef = "CreateObject\([`"'](?<progid>[\w\.]+)[`"']\)"

# Use Select String to do multiple macthes in a string, pick out the group value which is denoted in the regex by \w+.\w+
$progids = select-string -inputObject $source -Pattern $regexComRef -AllMatches | % { $_.Matches } | % { $_.Groups['progid'].Value }




foreach ( $progid in $progids ){
    
    $type = [System.Type]::GetTypeFromProgID($progid, $false)
    if ( $type -ne $null ) {
        write-host "Guid:" $type.GUID $type.lateBound
    }
    else {
        write-host "Missing Type:" $progid
    }

}
# print the progids out


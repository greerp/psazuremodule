remove-module Invoke-SQL
remove-module Get-ComRefs
import-module .\Invoke-SQL.ps1
Import-Module .\Get-ComRefs.ps1 

#(Invoke-SQL -dataSource "LLUCSAPP245P" -database "OfficeScanDB" `
#    -username "lsa_rbc_office2010" -password "Welcome001" `
#    -sqlCommand "select top 10 fileid, filepath+filename as filepath from euc_files where basefolder='Malaysia' and Loc `> 0 ")[0].Rows|
#%{@{id=$_.fileid;filepath=$_.filepath}}|
#Get-ComRefs|
#%{write-host "File:" $_.fileid ", Loc:" $_.Loc ", Hash:" $_.Codehash "}


#(Invoke-SQL -dataSource "LLUCSAPP245P" -database "OfficeScanDB" `
#    -username "lsa_rbc_office2010" -password "Welcome001" `
#    -sqlCommand "select top 10 fileid, filepath+filename as filepath from euc_files where basefolder='Malaysia' and Loc `> 0 ")[0].Rows|
#%{@{id=$_.fileid;filepath=$_.filepath}}|
#Get-ComRefs|
#%{ @{id=$_.fileid;loc=$_.Loc;hash=$_.CodeHash}|export-csv -path "output.csv" -append -NoTypeInformation -force}



$sqlQuery = "select fileid, filepath+filename as filepath from euc_files where basefolder='Malaysia' and Loc `> 0 "
$files = (Invoke-SQL -dataSource "LLUCSAPP245P" -database "OfficeScanDB" -username "lsa_rbc_office2010" -password "Welcome001" -sqlCommand $sqlQuery)[0].Rows

$result = $files|%{@{id=$_.fileid;filepath=$_.filepath}}|Get-ComRefs -verbose


$outfile = "kl-hashes.csv"
del $outfile -Force
foreach ( $item in $result ){
    New-Object psobject -Property @{
        FileId     = $item.fileid
        Loc        = $item.Loc
        Contains   = $item.CodeHash
        Comment    = $item.Comment } | export-csv -Path $outfile -append -NoTypeInformation -force
}    


foreach ( $item in $result ){
    New-Object psobject -Property @{
        FileId     = $item.fileid
        Loc        = $item.Loc
        Contains   = $item.CodeHash
        Comment    = $item.Comment } | export-csv -Path $outfile -append -NoTypeInformation -force
}  
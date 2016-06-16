<#
  .SYNOPSIS
  Gets a ProgID from a Clsid
  serves as a example how to call some embedded c# code from within a PS script

  .PARAMETER input
  Pipeline Strings with GUID string 
  

  .Example 1
  remove-module Get-TypeLib
  import-module .\Get-TypeLib.ps1
  $guids = import-csv -path mypath\myfile.csv
  $guids|%{@($_.guid)}|Get-TypeLib

  .Notes

#>
$Assem = ( 
    "system, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089, processorArchitecture=MSIL"
) 

$Source = @” 
    using System; 
    using System.Diagnostics;
    using System.Runtime.InteropServices;

    namespace com.RedPixie.euc 
    { 

        public class ComUtil 
        {
            [DllImport("ole32.dll", CharSet=CharSet.Unicode, PreserveSig=false)]
            static extern string ProgIDFromCLSID([In()]ref Guid clsid);

            public static string GetProgId(Guid guid) {
                return ProgIDFromCLSID(ref guid);
            } 
        }
    }
"@

function Get-ProgId($clsId) {
    $ex=""
    try { 
        $clsIdGuid = [Guid]$clsId
        $result = [com.RedPixie.euc.ComUtil]::GetProgId($clsIdGuid)
    }
    catch {
        $ex = $_
        $result = ""
    }
    return @{progid=$result;exception=$ex}
}


Add-Type -ReferencedAssemblies $Assem -TypeDefinition $Source -Language CSharp  
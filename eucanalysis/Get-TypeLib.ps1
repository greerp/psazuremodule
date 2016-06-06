function Get-TypeLib {
    param (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)] $value

    )

    begin {

        ###########################################################################
        function LoadMcdfLibrary {
            if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
                Write-Host "Loading MSOVBA Library"
            }
            [void][Reflection.Assembly]::LoadFile("${PSScriptRoot}\euclib.dll")

        }
    
        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-Host "Initialising...."
        }

        LoadMcdfLibrary

    }

    process {
        $guid = [Guid]$value;
        $count++;
        $result = [System.Collections.arrayList]@()

        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent){
            Write-host  "${count}: Reading typelib for  " $guid
        }

        $paths = [com.redpixie.euc.mcdf.PsWrapper]::GetLibraryForTypeLibGuid($guid)
        foreach ($path in $paths){
            if ( test-Path $path ){
                $result.Add(@{path=$path; IsBroken=$false})
            }
            else {
                $result.Add(@{path=$path; IsBroken=$true})
            }
        }
        
        $result
    }

    end {
    }
}


try {
    . .\authentication\Authenticate.ps1 -accountname "paul@greer.uk.com"
}
catch {

    Write-Host "An Error Occured in: " $_.InvocationInfo.ScriptName
    Write-Host "The Error was at line:" $_.InvocationInfo.ScriptLineNumber
    Write-Host "The actual Error is :" $_.Exception.Message

    # Might was to write error to the error stream using Write-Error
    write-Error "An Error Occured"

    # return Error to bamboo
    exit 1




}

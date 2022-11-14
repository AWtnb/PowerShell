
<# ==============================

cmdlets for searching file and folder

                encoding: utf8bom
============================== #>

function lsf {
    [OutputType("System.IO.FileInfo")]
    param()
    return Get-ChildItem -File $args
}

class PsCurrentDirExtension : System.Management.Automation.IValidateSetValuesGenerator {
    [string[]] GetValidValues() {
        return $((Get-ChildItem -File).Extension)
    }
}
function Get-FileByExtension {
    param (
        [ValidateSet([PsCurrentDirExtension])][string]$extension
    )
    return $(Get-ChildItem -File | Where-Object Extension -eq $extension)
}
Set-Alias lsx Get-FileByExtension


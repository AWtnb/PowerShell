
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

function nameBeginsWith {
    [OutputType("System.IO.FileInfo", "System.IO.DirectoryInfo")]
    param ([string]$name, [switch]$Recurce, [switch]$file, [switch]$directory)
    $filtered = @($input | Where-Object Name -Like "$name*")
    if ($filtered.Count) {
        return $filtered
    }
    return $(Get-ChildItem "$name*" -Recurse:$Recurce -Directory:$directory -File:$file)
}
function nameContains {
    [OutputType("System.IO.FileInfo", "System.IO.DirectoryInfo")]
    param ([string]$name, [switch]$Recurce, [switch]$file, [switch]$directory)
    $filtered = @($input | Where-Object Name -Like "*$name*")
    if ($filtered.Count) {
        return $filtered
    }
    return $(Get-ChildItem "*$name*" -Recurse:$Recurce -Directory:$directory -File:$file)
}
function nameEndsWith {
    [OutputType("System.IO.FileInfo", "System.IO.DirectoryInfo")]
    param ([string]$name, [switch]$Recurce, [switch]$file, [switch]$directory)
    $filtered = @($input | Where-Object BaseName -Like "*$name")
    if ($filtered.Count) {
        return $filtered
    }
    return $(Get-ChildItem -Recurse:$Recurce -Directory:$directory -File:$file | Where-Object BaseName -Like "*$name")
}


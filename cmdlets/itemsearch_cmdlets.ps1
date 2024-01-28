
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
    [OutputType("System.IO.FileInfo")]
    param (
        [ValidateSet([PsCurrentDirExtension])][string]$extension
    )
    return $(Get-ChildItem -File | Where-Object Extension -eq $extension)
}
Set-Alias lsx Get-FileByExtension

function lsc {
    $clip = (Get-Clipboard | Select-Object -First 1) -replace '"'
    if (Test-Path $clip -PathType Container) {
        Get-ChildItem -LiteralPath $clip | Write-Output
    }
    else {
        "invalid-path!" | Write-Host -ForegroundColor Magenta
    }
}


class Repos {
    [System.IO.DirectoryInfo[]]$repoDirs
    Repos([string[]]$paths) {
        $this.repoDirs = $paths | ForEach-Object {Get-Item $_} | Where-Object {-not $_.Extension} | Where-Object {Get-ChildItem $_ ".git" -Force}
    }

    [System.IO.DirectoryInfo[]]GetRemotes() {
        return $($this.repoDirs | Where-Object {
            $p = $_.Fullname
            return $(Get-Item "$p\.git\config" | Get-Content | Select-String -Pattern "^\[remote .+\]")
        })
    }

    [System.IO.DirectoryInfo[]]GetLocals() {
        return $($this.repoDirs | Where-Object {
            $p = $_.Fullname
            return -not $(Get-Item "$p\.git\config" | Get-Content | Select-String -Pattern "^\[remote .+\]")
        })
    }
}

function Find-RemoteRepository {
    $items = $input | ForEach-Object {Get-Item $_}
    return [Repos]::new($items).GetRemotes()
}

function Find-LocalRepository {
    $items = $input | ForEach-Object {Get-Item $_}
    return [Repos]::new($items).GetLocals()
}
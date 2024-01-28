
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
    [scriptblock]$checkBlock

    Repos([string[]]$paths) {
        $this.repoDirs = $paths | ForEach-Object {Get-Item $_} | Where-Object {-not $_.Extension} | Where-Object {$_ | Get-ChildItem -Filter ".git" -Force}
        $this.checkBlock = {
            param($path)
            $grep = (Get-Content -Path "$path\.git\config" | Select-String -Pattern "^\[remote .+\]")
            return $grep.Matches.Count -gt 0
        }
    }

    [System.IO.DirectoryInfo[]]GetRemotes() {
        return $($this.repoDirs | Where-Object {
            $p = $_.Fullname
            return $this.checkBlock.Invoke($p)
        })
    }

    [System.IO.DirectoryInfo[]]GetLocals() {
        return $($this.repoDirs | Where-Object {
            $p = $_.Fullname
            return -not $this.checkBlock.Invoke($p)
        })
    }
}

function Find-RemoteRepository {
    $paths = $input | ForEach-Object {Get-Item $_.Fullname}
    if ($paths.Length -lt 1) {
        $paths = (Get-ChildItem -Directory).FullName
    }
    return [Repos]::new($paths).GetRemotes()
}

function Find-LocalRepository {
    $paths = $input | ForEach-Object {Get-Item $_.Fullname}
    if ($paths.Length -lt 1) {
        $paths = (Get-ChildItem -Directory).FullName
    }
    return [Repos]::new($paths).GetLocals()
}
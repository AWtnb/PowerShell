
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


class Repo {
    [string]$path
    Repo([string]$path) {
        $this.path = $path
    }

    [bool]IsRemote() {
        if (-not (Test-Path $this.path -PathType Container)) {
            return $false
        }
        $g = Get-Item $this.path | Get-ChildItem -Filter ".git" -Force
        if (-not $g) {
            return $false
        }
        $grep = ($this.path | Join-Path -ChildPath ".git\config" | Get-Item | Select-String -Pattern "^\[remote .+\]")
        return $grep.Matches.Count -gt 0
    }
}

function Find-RemoteRepository {
    $paths = $input | ForEach-Object {Get-Item $_.Fullname}
    if ($paths.Count -lt 1) {
        $paths = (Get-ChildItem -Directory).FullName
    }
    return $paths | Where-Object { [Repo]::new($_).IsRemote() } | ForEach-Object { Get-Item $_ }
}

function Find-LocalRepository {
    $paths = $input | ForEach-Object {Get-Item $_.Fullname}
    if ($paths.Count -lt 1) {
        $paths = (Get-ChildItem -Directory).FullName
    }
    return $paths | Where-Object { -not [Repo]::new($_).IsRemote() } | ForEach-Object { Get-Item $_ }
}
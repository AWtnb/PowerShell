
<# ==============================

interactive filter

                encoding: utf8bom
============================== #>

# Both `[System.Console]::OutputEncoding` and `$OutputEncoding` must be UTF-8 to use fzf.exe


function Get-DefinitionOfCommand {
    $cmdletName = @(Get-Command -CommandType Function).Where({-not $_.Source}).Where({$_.Name -notmatch ":"}).Name | fzf.exe --no-color
    if ($cmdletName) {
        (Get-Command -Name $cmdletName).Definition | bat.exe --% --language=powershell
    }
}
Set-PSReadLineKeyHandler -Key "alt+f,d" -BriefDescription "fuzzyDefinition" -LongDescription "fuzzyDefinition" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::BeginningOfLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("<#SKIPHISTORY#> Get-DefinitionOfCommand #")
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}

class PSAvailable {
    [string]$cmdletsDir
    [System.IO.FileInfo[]]$files = @()
    [System.Collections.ArrayList]$sources

    PSAvailable() {
        $this.cmdletsDir = $Global:Profile | Split-Path -Parent | Join-Path -ChildPath "cmdlets"
        $d = Get-Item $this.cmdletsDir
        if ($d.LinkType -eq "Junction") {
            $this.cmdletsDir = $d.Target
        }
        $this.files +=  @($this.cmdletsDir | Get-ChildItem -File -Filter "*.ps1")
        $this.sources = New-Object System.Collections.ArrayList
    }

    SetData() {
        $this.SetFuncs()
        $this.SetClasses()
        $this.SetFiles()
    }

    Register([string]$name, [string]$path, [int]$linenum) {
        $this.sources.Add(
            [PSCustomObject]@{
                "name"    = $name;
                "path"    = $path;
                "lineNum" = $linenum
            }
        ) > $null
    }

    SetFuncs() {
        $this.files | Select-String -Pattern "^function" | ForEach-Object {
            $n = $_.line -replace "^function *" -replace "[ \(].*$"
            $this.Register($n, $_.Path, $_.LineNumber)
        }
    }

    SetClasses() {
        $this.files | Select-String -Pattern "^ *class" | ForEach-Object {
            $n = $_.line.trim() -replace " *{"
            $this.Register($n, $_.Path, $_.LineNumber)
        }
    }

    SetFiles() {
        $this.files | ForEach-Object {
            $n = "PS1:$($_.Basename)"
            $this.Register($n, $_.FullName, 1)
        }
        $p = $this.cmdletsDir | Split-Path -Parent | Join-Path -ChildPath "profile.ps1"
        $this.Register("PS1:Profile", $p, 1)
    }

    static [string[]] getCommands() {
        return @(Get-Command -CommandType Alias, Cmdlet, Function).Where({$_.Name -notmatch ":"}).Name
    }

}


Set-PSReadLineKeyHandler -Key "alt+f,spacebar" -BriefDescription "fuzzyfind-command" -LongDescription "search-cmdlets-with-fuzzyfinder" -ScriptBlock {
    $command = [PSAvailable]::getCommands() | fzf.exe --no-color
    if ($command) {
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert("$command ")
    }
    else {
        [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
    }
}


Set-PSReadLineKeyHandler -Key "alt+f,e" -BriefDescription "fuzzyEdit-customCmdlets" -LongDescription "fuzzyEdit-customCmdlets" -ScriptBlock {
    $c = [PSAvailable]::new()
    $c.SetData()
    $src = $c.sources
    $filtered = $src.Name | fzf.exe --no-color
    if ($filtered) {
        $selected = $src | Where-Object name -eq $filtered | Select-Object -First 1
        $wd = $c.cmdletsDir | Split-Path -Parent
        'code -g "{0}:{1}" "{2}"' -f $selected.path, $selected.lineNum, $wd | Invoke-Expression
    }
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}

function Invoke-DayPicker {
    param (
        [string]$start = ""
        ,[int]$y = 0
        ,[int]$m = 0
        ,[int]$d = 0
        ,[int]$span = 365
        ,[switch]$weekday
    )
    $opt = ($start)? @("--start", $start) : @()
    $opt = $opt + @("--year", $y, "--month", $m, "--day", $d, "--span", $span)
    if ($weekday) {
        $opt += "--weekday"
    }
    & fuzzy-daypick.exe $opt | Write-Output
}

function hinagata {
    $d = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\templates"
    if (-not (Test-Path $d)) {
        return
    }
    $n = Get-ChildItem -Path $d -Name -File | fzf.exe --layout=reverse --height=50% --no-color
    if (-not $n) {
        return
    }
    $p = $d | Join-Path -ChildPath $n
    $c = Get-Content -Path $p -Raw
    $param = [System.Web.HttpUtility]::UrlEncode($c)
    Start-Process ("https://awtnb.github.io/hinagata/?template={0}" -f $param)
}

function ghRemote {
    try {
        gh.exe repo list --json name --jq ".[] | .name" --limit 200 | fzf.exe --no-color --multi --layout=reverse --height=50% | Set-Variable -Name selected
    }
    catch {
        return
    }
    if ($selected.Count -lt 1) {
        return
    }
    $selected | Write-Host
    $clone = (Read-Host -Prompt "Clone? (y/N)") -eq "y"
    $selected | ForEach-Object {
        $repoName = $_
        if ($clone) {
            $url = "https://github.com/AWtnb/{0}.git" -f $repoName
            git clone $url
            if ($LASTEXITCODE -eq 0) {
                $cmd = Get-Command code -ErrorAction SilentlyContinue
                if ($cmd) {
                    Start-Process code -ArgumentList ($PWD.Path | Join-Path -ChildPath $repoName) -NoNewWindow
                }
            }
        }
        else {
            $url = "https://github.com/AWtnb/" + $repoName
            Start-Process $url
        }
    }
}

function Get-EjectableDrive {
    [OutputType([System.IO.DriveInfo])]
    param()
    return [System.IO.DriveInfo]::GetDrives() | Where-Object {$_.DriveType -eq [System.IO.DriveType]::Removable}
}

function Invoke-DriveEject {
    try {
        Get-Command fzf.exe -ErrorAction Stop > $null
    }
    catch {
        Write-Host "fzf.exe not found!"
        return
    }
    $drives = Get-EjectableDrive | ForEach-Object {
        if ($_.VolumeLabel) {
            return "{0} ({1})" -f $_.Name, $_.VolumeLabel
        }
        return $_.Name
    }
    if ($drives.Count -lt 1) {
        "No ejectable drive found" | Write-Host
        return
    }
    $d = $drives | fzf.exe --no-color --layout=reverse --height=50%
    if (-not $d) {
        return
    }
    $sh = New-Object -comObject Shell.Application
    $name = $d.Substring(0, $d.IndexOf("\"))
    $sh.NameSpace(17).ParseName($name).InvokeVerb("Eject")
}

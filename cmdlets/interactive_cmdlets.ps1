
<# ==============================

interactive filter

                encoding: utf8bom
============================== #>

# Both `[System.Console]::OutputEncoding` and `$OutputEncoding` must be UTF-8 to use fzf.exe


function Get-DefinitionOfCommand {
    $cmdletName = @(Get-Command -CommandType Function).Where({-not $_.Source}).Where({$_.Name -notmatch ":"}).Name | fzf.exe
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
    [string]$profPath
    [System.IO.FileInfo[]]$files = @()
    [System.Collections.ArrayList]$sources

    PSAvailable() {
        $this.profPath = $global:PROFILE
        $this.files = @(Get-Item -LiteralPath $this.profPath)
        $this.files +=  @($this.profPath | Split-Path -Parent | Join-Path -ChildPath "cmdlets" | Get-ChildItem -File -Filter "*.ps1")
        $this.sources = New-Object System.Collections.ArrayList
    }

    SetData() {
        $this.SetFuncs()
        $this.SetClasses()
        $this.SetPyCodes()
        $this.SetFiles()
    }

    SetFuncs() {
        $this.files | Select-String -Pattern "^function" | ForEach-Object {
            $this.sources.Add(
                [PSCustomObject]@{
                    "name"    = ($_.line -replace "^function *" -replace "[ \(].*$");
                    "path"    = $_.Path;
                    "lineNum" = $_.LineNumber;
                }
            ) > $null
        }
    }

    SetClasses() {
        $this.files | Select-String -Pattern "^ *class" | ForEach-Object {
            $this.sources.Add(
                [PSCustomObject]@{
                    "name"    = ($_.line.trim() -replace " *{");
                    "path"    = $_.Path;
                    "lineNum" = $_.LineNumber;
                }
            ) >$null
        }
    }

    SetPyCodes() {
        $pyDir = $this.profPath | Split-Path -Parent | Join-Path -ChildPath "cmdlets" | Join-Path -ChildPath "python"
        @($pyDir | Get-ChildItem -File -Filter "*.py" -Recurse) | ForEach-Object {
            $rel = [System.IO.Path]::GetRelativePath(($pyDir | Split-Path -Parent), $_.Fullname)
            $this.sources.Add(
                [PSCustomObject]@{
                    "name"    = $rel;
                    "path"    = $_.Fullname;
                    "lineNum" = 1;
                }
            ) >$null
        }
    }

    SetFiles() {
        $this.files | ForEach-Object {
            $this.sources.Add(
                [PSCustomObject]@{
                    "name"    = "PS1:$($_.Basename)";
                    "path"    = $_.Fullname;
                    "lineNum" = 1;
                }
            ) >$null
        }
        $this.sources.Add(
            [PSCustomObject]@{
                "name"    = "mdLess";
                "path"    = ($this.profPath | Split-Path -Parent | Join-Path -ChildPath "cmdlets\python\markdown\markdown.less");
                "lineNum" = 1;
            }
        ) >$null
        $this.sources.Add(
            [PSCustomObject]@{
                "name"    = "PS1:PROFILE";
                "path"    = $this.profPath;
                "lineNum" = 1;
            }
        ) >$null
    }

    static [string[]] getCommands() {
        return @(Get-Command -CommandType Alias, Cmdlet, Function).Where({$_.Name -notmatch ":"}).Name
    }

}


Set-PSReadLineKeyHandler -Key "alt+f,spacebar","ctrl+shift+spacebar" -BriefDescription "fuzzyfind-command" -LongDescription "search-cmdlets-with-fuzzyfinder" -ScriptBlock {
    $command = [PSAvailable]::getCommands() | fzf.exe
    if ($command) {
        [Microsoft.PowerShell.PSConsoleReadLine]::Insert("$command ")
    }
    else {
        [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
    }
}


Set-PSReadLineKeyHandler -Key "alt+f,e" -BriefDescription "fuzzyEdit-customCmdlets" -LongDescription "fuzzyEdit-customCmdlets" -ScriptBlock {
    $c = [PSAvailable]::new()
    $c.SetData()
    $src = $c.sources
    $filtered = $src.Name | fzf.exe
    if ($filtered) {
        $selected = $src | Where-Object name -eq $filtered | Select-Object -First 1
        $wd = $c.profPath | Split-Path -Parent
        'code -g "{0}:{1}" "{2}"' -f $selected.path, $selected.lineNum, $wd | Invoke-Expression
    }
    [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
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
    & ($env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\fuzzy-daypick.exe") $opt | Write-Output
}
function Invoke-DayPickerClipper {
    param (
        [string]$start = ""
        ,[int]$y = 0
        ,[int]$m = 0
        ,[int]$d = 0
        ,[int]$span = 365
        ,[switch]$weekday
    )
    $days = Invoke-DayPicker -start $start -y $y -m $m -d $d -span $span -weekday:$weekday
    if ($days -and $LASTEXITCODE -eq 0) {
        $days | Set-Clipboard
        "coplied:" | Write-Host -ForegroundColor Blue
        $days | Write-Host
        [System.Windows.Forms.SendKeys]::SendWait("%{Tab}")
    }
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+d" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("Invoke-DayPickerClipper -weekday")
}


Set-PSReadLineKeyHandler -Key "ctrl+alt+z","alt+z" -ScriptBlock {
    param($key, $arg)
    $dataPath = ($env:USERPROFILE | Join-Path -ChildPath "Personal\launch.yaml")
    $opt = @("--src", $dataPath, "--filer", $env:TABLACUS_PATH, "--exclude=_obsolete,node_modules")
    if ($key.Modifiers -band [System.ConsoleModifiers]::Control) {
        $opt += "--all"
    }
    & ($env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\moko.exe") $opt
}

function Invoke-RDriveDatabase {
    param (
        [string]$name
    )
    $dirs = Get-ChildItem "R:" -Directory
    if ($name) {
        $reg = [regex]$name
        $dirs = @($dirs).Where({$reg.IsMatch($_.name)})
    }
    if ($dirs.count -gt 1) {
        $filtered = $dirs.name | fzf.exe
        if (-not $filtered) {
            return
        }
        "R:" | Join-Path -ChildPath $filtered | Invoke-Item
    }
    else {
        $dirs[0].fullname | Invoke-Item
    }
}

function hinagata {
    $d = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\templates"
    if (-not (Test-Path $d)) {
        return
    }
    $n = Get-ChildItem -Path $d -Name | fzf.exe
    if (-not $n) {
        return
    }
    $p = $d | Join-Path -ChildPath $n
    Get-Item -Path $p | Get-Content | Set-Clipboard
    Start-Process "https://awtnb.github.io/hinagata/"
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+h" -ScriptBlock {
    hinagata
    [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
}

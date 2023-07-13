
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

    static [string[]]$commands = @(Get-Command -CommandType Alias, Cmdlet, Function).Where({$_.Name -notmatch ":"}).Name

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
                    "name" = ($_.line -replace "^function *" -replace "[ \(].*$");
                    "path" = $_.Path;
                    "lineNum" = $_.LineNumber;
                }
            ) > $null
        }
    }

    SetClasses() {
        $this.files | Select-String -Pattern "^ *class" | ForEach-Object {
            $this.sources.Add(
                [PSCustomObject]@{
                    "name" = ($_.line.trim() -replace " *{");
                    "path" = $_.Path;
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
                   "name" = $rel;
                   "path" = $_.Fullname;
                   "lineNum" = 1;
               }
            ) >$null
        }
    }

    SetFiles() {
        $this.files | ForEach-Object {
            $this.sources.Add(
                [PSCustomObject]@{
                   "name" = "PS1:$($_.Basename)";
                   "path" = $_.Fullname;
                   "lineNum" = 1;
               }
            ) >$null
        }
        $this.sources.Add(
            [PSCustomObject]@{
               "name" = "mdLess";
               "path" = ($this.profPath | Split-Path -Parent | Join-Path -ChildPath "cmdlets\python\markdown\markdown.less");
               "lineNum" = 1;
           }
        ) >$null
        $this.sources.Add(
            [PSCustomObject]@{
               "name" = "PS1:PROFILE";
               "path" = $this.profPath;
               "lineNum" = 1;
           }
        ) >$null
    }

}


Set-PSReadLineKeyHandler -Key "alt+f,spacebar","ctrl+shift+spacebar" -BriefDescription "fuzzyfind-command" -LongDescription "search-cmdlets-with-fuzzyfinder" -ScriptBlock {
    $command = [PSAvailable]::commands | fzf.exe
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


function Invoke-MokoLauncher {
    param (
        [switch]$all
    )
    $dataPath = "C:\Personal\launch.yaml"
    if (-not (Test-Path $dataPath)) {
        "cannnot find '{0}'" -f $dataPath | Write-Host -ForegroundColor Red
        return
    }
    $opt = @("--src", $dataPath, "--filer", $env:TABLACUS_PATH, "--exclude=_obsolete,node_modules")
    if ($all) {
        $opt += "--all"
    }
    & "C:\Personal\tools\bin\moko.exe" $opt
    if ($LASTEXITCODE -eq 0) {
        Hide-ConsoleWindow
    }
}
Set-Alias moko Invoke-MokoLauncher

Set-PSReadLineKeyHandler -Key "alt+z" -ScriptBlock {
    Invoke-MokoLauncher
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+z" -ScriptBlock {
    Invoke-MokoLauncher -all
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
}

function hinagata {
    $templateDir = "C:\Personal\tools\templates"
    if (Test-Path -Path $templateDir -PathType Container) {
        $names = ($templateDir | Get-ChildItem -Filter "*.txt").Name
        $selected = $names | mokof -ascii
        if ($selected) {
            $item = $templateDir | Join-Path -ChildPath $selected | Get-Item
            $item | Get-Content | Set-Clipboard
            "COPIED: '{0}'" -f $item.Name | Write-Host -ForegroundColor Yellow
            Start-Process "https://awtnb.github.io/hinagata/"
        }
    }
    else {
        "Cannot find templates..." | Write-Host -ForegroundColor Magenta
    }
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

<# ==============================

interactive filter

                encoding: utf8bom
============================== #>

function mokof {
    param(
        [switch]$ascii
    )
    $exePath = "C:\Personal\tools\bin\mokof.exe"
    if ($ascii) {
        $input | & $exePath | Write-Output
    }
    else {
        $byte = $input | & $exePath "--bytearr"
        if ($LASTEXITCODE -eq 0) {
            [System.Text.Encoding]::UTF8.GetString($byte) | Write-Output
        }
    }
}


function Get-DefinitionOfCommand {
    $cmdletName = @(Get-Command -CommandType Function).Where({-not $_.Source}).Where({$_.Name -notmatch ":"}).Name | mokof -ascii
    if ($cmdletName) {
        (Get-Command -Name $cmdletName).Definition | bat.exe --% --language=powershell
    }
}
Set-PSReadLineKeyHandler -Key "alt+f,d" -BriefDescription "fuzzyDefinition" -LongDescription "fuzzyDefinition" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("<#SKIPHISTORY#> Get-DefinitionOfCommand")
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}

class PSAvailable {
    [string]$profPath
    [System.IO.FileInfo[]]$files = @()
    [System.Collections.ArrayList]$sources

    static [string[]]$commands = @(Get-Command -CommandType Alias, Cmdlet, Function).Where({$_.Name -notmatch ":"}).Name

    PSAvailable() {
        $this.profPath = $env:USERPROFILE | Join-Path -ChildPath "\Documents\PowerShell\Microsoft.PowerShell_profile.ps1"
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
    $command = [PSAvailable]::commands | mokof -ascii
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
    $filtered = $src.Name | mokof -ascii
    if ($filtered) {
        $selected = $src | Where-Object name -eq $filtered | Select-Object -First 1
        # $wd = $env:USERPROFILE | Join-Path -ChildPath "Dropbox\develop\app_config\PowerShell\PowerShell"
        $wd = $c.profPath | Split-Path -Parent
        'code -g "{0}:{1}" "{2}"' -f $selected.path, $selected.lineNum, $wd | Invoke-Expression
    }
    [Microsoft.PowerShell.PSConsoleReadLine]::InvokePrompt()
}


function psMoko {
    param (
        [switch]$all
        ,[string]$exclude = "_obsolete,node_modules"
    )
    $dataPath = "C:\Personal\launch.yaml"
    if (-not (Test-Path $dataPath)) {
        "cannnot find '{0}'" -f $dataPath | Write-Host -ForegroundColor Red
        return
    }
    $exePath = "C:\Personal\tools\bin\moko.exe"
    $filerPath = "C:\Users\{0}\Dropbox\portable_apps\tablacus\TE64.exe" -f $env:USERNAME
    $opt = @("--src", $dataPath, "--filer", $filerPath, "--exclude", $exclude)
    if ($all) {
        $opt += "--all"
    }
    & $exePath $opt
    if ($LASTEXITCODE -eq 0) {
        Hide-ConsoleWindow
    }
}
Set-Alias z psMoko
Set-PSReadLineKeyHandler -Key "alt+z" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert('<#SKIPHISTORY#> z')
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+z" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert('<#SKIPHISTORY#> z -all')
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
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

# function Invoke-DesktopItem ([switch]$file) {
#     $desktop = "C:\Users\{0}\Desktop" -f $env:USERNAME
#     $items = @(Get-ChildItem $desktop -Name -File:$file)
#     if ($items) {
#         $filtered = $items | mokof
#         if ($filtered) {
#             return $($desktop | Join-Path -ChildPath $filtered | Get-Item)
#         }
#     }
# }

# function ide {
#     $item = Invoke-DesktopItem
#     if ($item) {
#         Invoke-Item $item
#         Hide-ConsoleWindow
#     }
# }
# function idc {
#     $item = Invoke-DesktopItem -file
#     if ($item) {
#         [System.Windows.Forms.Clipboard]::SetFileDropList($item)
#     }
# }
# function idx {
#     $item = Invoke-DesktopItem -file
#     if (-not $item) {
#         return
#     }
#     $fullname = $item.Fullname
#     try {
#         $dataObj = New-Object System.Windows.Forms.DataObject
#         $dataObj.SetFileDropList($fullname)
#         $byteStream = [byte[]](([System.Windows.Forms.DragDropEffects]::Move -as [byte]), 0, 0, 0)
#         $memoryStream = New-Object System.IO.MemoryStream
#         $memoryStream.Write($byteStream)
#         $dataObj.SetData("Preferred DropEffect", $memoryStream)
#         [System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)

#         Write-Host "CUT item on desktop: " -NoNewline
#         $color = ($item.GetType().Name -eq "DirectoryInfo")? "Yellow" : "Blue"
#         Write-Host $item.Name -ForegroundColor $color
#         Invoke-Taskview
#     }
#     catch {
#         Write-Host $_.Exception.Message -ForegroundColor Red
#     }
# }


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
        $filtered = $dirs.name | mokof
        if (-not $filtered) {
            return
        }
        "R:" | Join-Path -ChildPath $filtered | Invoke-Item
    }
    else {
        $dirs[0].fullname | Invoke-Item
    }
}
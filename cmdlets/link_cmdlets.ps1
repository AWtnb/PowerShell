
<# ==============================

cmdlets for treating symlink / junction / shortcut

                encoding: utf8bom
============================== #>

function Get-Shortcut {
    <#
        .EXAMPLE
        ls | Get-Shortcut
    #>
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -in @(".lnk", ".url")) {
            $shell = New-Object -ComObject WScript.Shell
            $shell.CreateShortcut($fileObj.Fullname) | Write-Output
        }
    }
    end {}
}

function Set-ShortCutHotkey {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
        ,[string]$hotkey = ""
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -in @(".lnk", ".url")) {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($fileObj.Fullname)
            if (-not $hotkey.Length) {
                $org = $shortcut.Hotkey
                $shortcut.Hotkey = ""
                $shortcut.Save()
                "cleared hotkey '{0}'!" -f $org | Write-Host
                return
            }
            if ($shortcut.Hotkey) {
                "this shortcut has hotkey: {0}" -f $shortcut.Hotkey | Write-Host
                $ask = Read-Host "overwrite? (y/n)"
                if ($ask -ne "y") {
                    return
                }
            }
            $shortcut.Hotkey = $hotkey
            $shortcut.Save()
            "set hotkey '{0}'!" -f $hotkey | Write-Host
        }
    }
    end {}
}

class PsLinker {
    [string]$srcPath
    [string]$srcName
    [string]$workDir
    [string]$linkPath
    PsLinker([string]$srcPath, [string]$workDir) {
        $src = Get-Item -LiteralPath $srcPath
        $this.srcPath = $src.FullName
        $this.srcName = $src.Name
        $this.workDir = ($workDir.Length -gt 0)? (Get-Item -LiteralPath $workDir).FullName : (Get-Location).Path
        $this.linkPath = $this.workDir | Join-Path -ChildPath $this.srcName
    }

    AskInvoke() {
        "'{0}' already exists on '{1}'!" -f $this.srcName, $this.workDir | Write-Host -ForegroundColor Red
        if ((Read-Host "open the directory? (y/n)") -eq "y") {
            Invoke-Item $this.workDir
        }
    }

    MakeSymbolicLink() {
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            "Need to run as ADMIN to make symlnk of '{0}'..." -f $this.srcPath | Write-Host -ForegroundColor Red
            return
        }
        if (Test-Path $this.linkPath) {
            $this.AskInvoke()
            return
        }
        try {
            New-Item -Path $this.linkPath -Value $this.srcPath -ItemType SymbolicLink -ErrorAction Stop
        }
        catch {
            "failed to make new SymbolicLink '{0}'!" -f $this.linkPath | Write-Error
        }
    }

    MakeJunction() {
        if (Test-Path $this.linkPath) {
            $this.AskInvoke()
            return
        }
        try {
            New-Item -Path $this.linkPath -Value $this.srcPath -ItemType Junction -ErrorAction Stop
        }
        catch {
            "failed to make new Junction '{0}'!" -f $this.linkPath | Write-Error
        }
    }

    MakeShortcut() {
        $shortcutPath = $this.linkPath + ".lnk"
        if (Test-Path $shortcutPath) {
            $this.AskInvoke()
            return
        }
        $wsShell = New-Object -ComObject WScript.Shell
        $shortcut = $WsShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $this.srcPath
        $shortcut.Save()
    }
}


function New-SymbolicLink {
    param (
        [parameter(Mandatory)][string]$src
        ,[string]$linkLocation
    )
    $linker = [PsLinker]::New($src, $linkLocation)
    $linker.MakeSymbolicLink()
}

function New-SymlnkOnPersonalBin {
    param (
        [parameter(Mandatory)][string]$src
    )
    $linkLocation = "C:\Personal\tools\bin"
    if (-not (Test-Path $linkLocation)) {
        "'{0}' not exists!" -f $linkLocation | Write-Host -ForegroundColor Red
        return
    }
    $linker = [PsLinker]::New($src, $linkLocation)
    $linker.MakeSymbolicLink()
}

function New-Junction {
    param (
        [parameter(Mandatory)][string]$src
        ,[string]$junctionLocation
    )
    $linker = [PsLinker]::New($src, $junctionLocation)
    $linker.MakeJunction()
}

function New-ShortCut {
    param (
        [parameter(Mandatory)][string]$src
        ,[string]$shortcutPlace
    )
    $linker = [PsLinker]::New($src, $shortcutPlace)
    $linker.MakeShortcut()
}

function New-ShortCutOnStartup {
    param (
        [parameter(Mandatory)][string]$path
    )
    if (Test-Path $path) {
        $startup = $env:USERPROFILE | Join-Path -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        New-ShortCut -src $path -shortcutPlace $startup
    }
    else {
        "INVALID PATH! : '{0}'" -f $path | Write-Host -ForegroundColor Magenta
    }
}

function New-ShortCutOnMyDataSources {
    <#
        .EXAMPLE
        New-ShortCutOnMyDataSources "C:\Personal\hogehoge.txt" # => create lnk to "~\Documents\My Data Sources"
    #>
    param (
        [string]$path
    )
    if (Test-Path $path) {
        $myDataSource = "C:\Users\{0}\Documents\My Data Sources" -f $env:USERNAME
        New-ShortCut -src $path -shortcutPlace $myDataSource
        "created shortcut '{0}' on 'MY DATA SOURCE'!" -f ($path | Split-Path -Leaf) | Write-Host -ForegroundColor Green
    }
    else {
        "INVALID PATH! : '{0}'" -f $path | Write-Host -ForegroundColor Magenta
    }
}


function Set-ShortcutFiler {
    param (
        $filerPath = "$($env:USERPROFILE)\Dropbox\portable_apps\tablacus\TE64.exe"
    )
    begin {
        $shell  = New-Object -ComObject WScript.Shell
    }
    process {
        if ($_.Extension -ne ".lnk") {
            "{0} is not shortcut file!" -f $_.Name | Write-Host -ForegroundColor Magenta
            return
        }
        if ($shell.CreateShortcut($_).TargetPath -eq $filerPath) {
            "filer of {0} is already modified!" -f $_.Name | Write-Host -ForegroundColor Magenta
            return
        }
        $tmpLnk = $shell.CreateShortcut($_)
        $openDir = $tmpLnk.TargetPath
        if (-not (Test-Path $openDir -PathType Container)) {
            "{0} is not shortcut to folder!" -f $_.Name | Write-Host -ForegroundColor Magenta
            return
        }
        $tmpLnk.Arguments = $openDir
        $tmpLnk.TargetPath = $filerPath
        $tmpLnk.save()
    }
    end {
    }
}

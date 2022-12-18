
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



function Test-Admin {
    return (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
}

function New-SymbolicLink {
    param (
        [parameter(Mandatory)][string]$src
        ,[string]$linkLocation
    )
    $wd = ($linkLocation.Length -gt 0)? (Get-Item -LiteralPath $linkLocation).FullName : (Get-Location).Path
    $linkSrc = Get-Item -LiteralPath $src
    $linkPath = $wd | Join-Path -ChildPath $linkSrc.Name
    if (Test-Path $linkPath) {
        "'{0}' already exists!" -f $linkPath | Write-Host -ForegroundColor Red
        $ask = Read-Host "open the directory? (y/n)"
        if ($ask -eq "y") {
            Invoke-Item $wd
        }
        return
    }
    try {
        New-Item -Path $linkPath -Value $linkSrc.FullName -ItemType SymbolicLink -ErrorAction Stop
    }
    catch {
        "failed to make new SymbolicLink '{0}'!" -f $linkPath | Write-Error
    }
}

function New-Junction {
    param (
        [parameter(Mandatory)][string]$src
        ,[string]$junctionLocation
    )
    $wd = ($junctionLocation.Length -gt 0)? (Get-Item -LiteralPath $junctionLocation).FullName : (Get-Location).Path
    $linkSrc = Get-Item -LiteralPath $src
    $jctPath = $wd | Join-Path -ChildPath $linkSrc.Name
    if (Test-Path $jctPath) {
        "'{0}' already exists!" -f $jctPath | Write-Host -ForegroundColor Red
        $ask = Read-Host "open the directory? (y/n)"
        if ($ask -eq "y") {
            Invoke-Item $wd
        }
        return
    }
    try {
        New-Item -Path $jctPath -Value $linkSrc.FullName -ItemType Junction -ErrorAction Stop
    }
    catch {
        "failed to make new Junction '{0}'!" -f $jctPath | Write-Error
    }
}

function New-ShortCut {
    param (
        [parameter(Mandatory)][string]$pathToJump
        ,[string]$shortcutPlace
    )
    $linkSrc = Get-Item -LiteralPath $pathToJump
    $linkName = $linkSrc.BaseName + ".lnk"
    $wd = ($shortcutPlace.Length -gt 0)? (Get-Item -LiteralPath $shortcutPlace).FullName : (Get-Location).Path
    $shortcutPath = $wd | Join-Path -ChildPath $linkName
    if (Test-Path $shortcutPath) {
        "'{0}' already exists!" -f $shortcutPath | Write-Host -ForegroundColor Red
        $ask = Read-Host "open the directory? (y/n)"
        if ($ask -eq "y") {
            Invoke-Item $wd
        }
        return
    }
    $wsShell = New-Object -ComObject WScript.Shell
    $shortcut = $WsShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $linkSrc.FullName
    $shortcut.Save()
}

function New-ShortCutOnStartup {
    param (
        [parameter(Mandatory)][string]$path
    )
    if (Test-Path $path) {
        $startup = $env:USERPROFILE | Join-Path -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        New-ShortCut -pathToJump $path -shortcutPlace $startup
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
        New-ShortCut -pathToJump $path -shortcutPlace $myDataSource
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

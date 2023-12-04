
<# ==============================

cmdlets for renaming file or folder

                encoding: utf8bom
============================== #>

class BasenameReplaceEntry {
    [regex]$_reg
    [string]$_execColor = "Green"
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    [string]$_newName
    BasenameReplaceEntry([string]$fullName, [string]$curDir, [string]$from, [string]$to, [switch]$case) {
        $regOpt = ($case)? [System.Text.RegularExpressions.RegexOptions]::None : [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        $this._reg = [regex]::new($from, $regOpt)
        $this._fullName = $fullName
        $item = Get-Item $this._fullName
        $this._orgBaseName = $item.BaseName
        $this._extension = $item.Extension
        $this._relDirName = [System.IO.Path]::GetRelativePath($curDir, ($this._fullName | Split-Path -Parent))
        $this._newName = $this._reg.Replace($this._orgBaseName, $to) + $this._extension
    }

    [bool] isRenamable() {
        return -not (($this._orgBaseName + $this._extension) -ceq $this._newName)
    }

    [string] getFullname() {
        return $this._fullName
    }

    [string] getNewName() {
        return $this._newName
    }

    [string] _getMarkerdNewName() {
        return $global:PSStyle.Foreground.PSObject.Properties[$this._execColor].Value + $this._newName + $global:PSStyle.Reset
    }

    [string] _getMatchesMarkerdOrgName([bool]$execute) {
        $col = ($execute)? $this._execColor : "White"
        $ansi = $global:PSStyle.Background.PSObject.Properties[$col].Value + $global:PSStyle.Foreground.Black
        return $this._reg.Replace($this._orgBaseName, {
                param([System.Text.RegularExpressions.Match]$m)
                return $ansi + $m.Value + $global:PSStyle.Reset
            }) + $this._extension
    }

    [string] _getDimmedRelDir() {
        return $global:PSStyle.Foreground.BrightBlack + $this._relDirName + "\" + $global:PSStyle.Reset
    }

    [int] getLeftSideByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._orgBaseName + $this._extension)
    }

    [string] getFullMarkerdText([bool]$org, [bool]$execute) {
        if ($org) {
            return $this._getDimmedRelDir() + $this._getMatchesMarkerdOrgName($execute)
        }
        return $this._getDimmedRelDir() + $this._getMarkerdNewName()
    }

}

class BasenameReplacer {
    [int]$_leftBufferWidth = 0
    [BasenameReplaceEntry[]]$entries = @()
    BasenameReplacer() {}

    [void] setEntry([BasenameReplaceEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        if ($this._leftBufferWidth -lt $w) {
            $this._leftBufferWidth = $w
        }
    }

    [string] getFiller([int]$indent) {
        $rightPadding = [Math]::Max($this._leftBufferWidth - $indent, 0)
        $filler = " {0}=> " -f ("=" * $rightPadding)
        return $Global:PSStyle.Foreground.Yellow + $filler + $Global:PSStyle.Reset
    }

    [void] run([bool]$execute) {
        $this.entries | ForEach-Object {
            $left = $_.getFullMarkerdText($true, $execute)
            $indent = $_.getLeftSideByteLen()
            $filler = $this.getFiller($indent)
            $right = $_.getFullMarkerdText($false, $execute)
            $left + $filler + $right | Write-Host
            if (-not $execute) {
                return
            }
            $item = Get-Item -LiteralPath $_.getFullname()
            $newName = $_.getNewName()
            try {
                $item | Rename-Item -NewName $newName -ErrorAction Stop
            }
            catch {
                "same file '{0}' already exists!" -f $newName | Write-Host -ForegroundColor Magenta
            }
        }
    }

}

function Rename-ReplaceBasename {
    <#
        .EXAMPLE
        ls *.txt | repn "foo" "baa" -execute
    #>
    [OutputType([System.Void])]
    param (
        [string]$from
        ,[string]$to
        ,[switch]$case
        ,[switch]$execute
    )

    $replacer = [BasenameReplacer]::new()
    $cur = (Get-Location).Path
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [BasenameReplaceEntry]::new($_.Fullname, $cur, $from, $to, $case)
        if ($ent.isRenamable()) {
            $replacer.setEntry($ent)
        }
    }
    $replacer.run($execute)
}
Set-Alias repBN Rename-ReplaceBasename


class InsertRenameEntry {
    [int]$_pos
    [string]$_pre
    [string]$_suf
    [string]$_insert
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    InsertRenameEntry([string]$fullName, [string]$curDir, [string]$insert, [int]$pos) {
        $this._pos = $pos
        $this._fullName = $fullName
        $item = Get-Item $this._fullName
        $this._orgBaseName = $item.BaseName
        $this._extension = $item.Extension
        $this._relDirName = [System.IO.Path]::GetRelativePath($curDir, ($this._fullName | Split-Path -Parent))
        if ([Math]::Abs($this._pos) -gt $item.BaseName.Length + 1) {
            $this._pre = $item.Basename
            $this._insert = ""
            $this._suf = ""
            return
        }
        $this._insert = $insert
        if ($this._pos -eq 0) {
            $this._pre = ""
            $this._suf = $item.Basename
            return
        }
        if ($this._pos -lt 0) {
            $this._pre = ($item.BaseName).substring(0, ($item.BaseName).length + $this._pos + 1)
            $this._suf = ($item.BaseName).substring(   ($item.BaseName).length + $this._pos + 1)
            return
        }
        $this._pre = ($item.BaseName).substring(0, $this._pos)
        $this._suf = ($item.BaseName).substring($this._pos)
    }

    [string] getFullName() {
        return $this._fullName
    }

    [string] getNewName() {
        return $this._pre + $this._insert + $this._suf + $this._extension
    }

    [bool] isRenamable() {
        return -not (($this._orgBaseName + $this._extension) -ceq $this.getNewName())
    }

    [string] _getMarkerdNewName([bool]$execute) {
        $color = ($execute)? "Green" : "White"
        return $this._pre + `
            $Global:PSStyle.Foreground.Black + `
            $Global:PSStyle.Background.PSObject.Properties[$color].Value + `
            $this._insert + `
            $Global:PSStyle.Reset + `
            $this._suf + `
            $this._extension
    }

    [string] _getDimmedRelDir() {
        return $global:PSStyle.Foreground.BrightBlack + $this._relDirName + "\" + $global:PSStyle.Reset
    }

    [int] getLeftSideByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._pre)
    }

    [string] getFullMarkerdText([bool]$execute) {
        return $this._getDimmedRelDir() + $this._getMarkerdNewName($execute)
    }

}


class InsertRenamer {
    [int]$_leftBufferWidth = 0
    [InsertRenameEntry[]]$entries = @()
    InsertRenamer() {}

    [void] setEntry([InsertRenameEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        if ($this._leftBufferWidth -lt $w) {
            $this._leftBufferWidth = $w
        }
    }

    [string] getFiller([int]$leftSideLen) {
        $padding = [Math]::Max($this._leftBufferWidth - $leftSideLen, 0)
        return " " * $padding
    }

    [void] run($execute) {
        $this.entries | ForEach-Object {
            $indent = $_.getLeftSideByteLen()
            $filler = $this.getFiller($indent)
            $filler + $_.getFullMarkerdText($execute) | Write-Host
            if (-not $execute) {
                return
            }
            $item = Get-Item -LiteralPath $_.getFullname()
            $newName = $_.getNewName()
            try {
                $item | Rename-Item -NewName $newName -ErrorAction Stop
            }
            catch {
                "failed to rename as '{0}'!" -f $newName | Write-Host -ForegroundColor Magenta
            }
        }
    }
}


function Rename-Insert {
    <#
        .EXAMPLE
        ls * | Rename-Insert "hoge_" -execute
    #>
    [OutputType([System.Void])]
    param (
        [string]$insert
        ,[int]$position = -1
        ,[switch]$execute
    )

    $renamer = [InsertRenamer]::new()
    $cur = (Get-Location).Path
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [InsertRenameEntry]::new($_.Fullname, $cur, $insert, $position)
        if ($ent.isRenamable()) {
            $renamer.setEntry($ent)
        }
    }
    $renamer.run($execute)
}
Set-Alias rIns Rename-Insert


class IndexRenameEntry {
    [string]$_idx
    [string]$_pre
    [string]$_highlight
    [string]$_suf
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    IndexRenameEntry([string]$fullName, [string]$curDir, [string]$altName, [int]$idx, [int]$pad, [bool]$tail) {
        $this._idx = ($idx -as [string]).PadLeft($pad, "0")
        $this._fullName = $fullName
        $item = Get-Item $this._fullName
        $this._orgBaseName = $item.BaseName
        $this._extension = $item.Extension
        $this._relDirName = [System.IO.Path]::GetRelativePath($curDir, ($this._fullName | Split-Path -Parent))
        if ($altName.Length -gt 0) {
            $this._pre = ""
            $this._suf = $this._extension
            if ($tail) {
                $this._highlight = $altName + $this._idx
                return
            }
            $this._highlight = $this._idx + $altName
            return
        }
        if ($tail) {
            $this._pre = $this._orgBaseName
            $this._highlight = "_" + $this._idx
            $this._suf = $this._extension
            return
        }
        $this._pre = ""
        $this._highlight = $this._idx + "_"
        $this._suf = $this._orgBaseName + $this._extension
    }

    [string] getFullName() {
        return $this._fullName
    }

    [string] getNewName() {
        return $this._pre + $this._highlight + $this._suf
    }

    [bool] isRenamable() {
        return -not (($this._orgBaseName + $this._extension) -ceq $this.getNewName())
    }

    [string] _getMarkerdNewName([bool]$execute) {
        $color = ($execute)? "Green" : "White"
        return $this._pre + `
            $Global:PSStyle.Foreground.Black + `
            $Global:PSStyle.Background.PSObject.Properties[$color].Value + `
            $this._highlight + `
            $Global:PSStyle.Reset + `
            $this._suf
    }

    [string] _getDimmedRelDir() {
        return $global:PSStyle.Foreground.BrightBlack + $this._relDirName + "\" + $global:PSStyle.Reset
    }

    [int] getLeftSideByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._pre)
    }

    [string] getFullMarkerdText([bool]$execute) {
        return $this._getDimmedRelDir() + $this._getMarkerdNewName($execute)
    }

}

class IndexRenamer {
    [int]$_leftBufferWidth = 0
    [IndexRenameEntry[]]$entries = @()
    IndexRenamer() {}

    [void] setEntry([IndexRenameEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        if ($this._leftBufferWidth -lt $w) {
            $this._leftBufferWidth = $w
        }
    }

    [string] getFiller([int]$leftSideLen) {
        $padding = [Math]::Max($this._leftBufferWidth - $leftSideLen, 0)
        return " " * $padding
    }

    [void] run($execute) {
        $this.entries | ForEach-Object {
            $indent = $_.getLeftSideByteLen()
            $filler = $this.getFiller($indent)
            $filler + $_.getFullMarkerdText($execute) | Write-Host
            if (-not $execute) {
                return
            }
            $item = Get-Item -LiteralPath $_.getFullname()
            $newName = $_.getNewName()
            try {
                $item | Rename-Item -NewName $newName -ErrorAction Stop
            }
            catch {
                "failed to rename as '{0}'!" -f $newName | Write-Host -ForegroundColor Magenta
            }
        }
    }

}

function Rename-Index {
    <#
        .EXAMPLE
        ls * | Rename-Index
    #>
    [OutputType([System.Void])]
    param (
        [string]$altName
        ,[int]$start = 1
        ,[int]$pad = 2
        ,[int]$step = 1
        ,[int[]]$skips
        ,[switch]$tail
        ,[switch]$execute
    )

    $renamer = [IndexRenamer]::new()
    $cur = (Get-Location).Path
    $i = $start
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [IndexRenameEntry]::new($_.Fullname, $cur, $altName, $i, $pad, $tail)
        if ($ent.isRenamable()) {
            $renamer.setEntry($ent)
            $i += $step
            if ($skips.Length) {
                while ($i -in $skips) {
                    $i += $step
                }
            }
        }
    }
    $renamer.run($execute)
}
Set-Alias rInd Rename-Index

class NameReplaceEntry {
    [string]$_newName
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    NameReplaceEntry([string]$fullName, [string]$curDir, [string]$newName) {
        $this._newName = $newName
        $this._fullName = $fullName
        $item = Get-Item $this._fullName
        $this._orgBaseName = $item.BaseName
        $this._extension = $item.Extension
        $this._relDirName = [System.IO.Path]::GetRelativePath($curDir, ($this._fullName | Split-Path -Parent))
    }

    [bool] isRenamable() {
        return -not (($this._orgBaseName + $this._extension) -ceq $this._newName)
    }

    [string] getFullname() {
        return $this._fullName
    }

    [string] getNewName() {
        return $this._newName
    }

    [string] _getDimmedRelDir() {
        return $global:PSStyle.Foreground.BrightBlack + $this._relDirName + "\" + $global:PSStyle.Reset
    }

    [int] getLeftSideByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._orgBaseName + $this._extension)
    }

    [string] getFullMarkerdText([bool]$org, [bool]$execute) {
        $color = ($org)? "White" : "Green"
        $ansi = $global:PSStyle.Foreground.PSObject.Properties[$color].Value
        $n = ($org)? ($this._orgBaseName + $this._extension) : $this._newName
        return $this._getDimmedRelDir() + $ansi + $n + $global:PSStyle.Reset
    }
}

class NameReplacer {
    [int]$_leftBufferWidth = 0
    [NameReplaceEntry[]]$entries = @()
    NameReplacer() {}

    [void] setEntry([NameReplaceEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        if ($this._leftBufferWidth -lt $w) {
            $this._leftBufferWidth = $w
        }
    }

    [string] getFiller([int]$indent) {
        $rightPadding = [Math]::Max($this._leftBufferWidth - $indent, 0)
        $filler = " {0}=> " -f ("=" * $rightPadding)
        return $Global:PSStyle.Foreground.Yellow + $filler + $Global:PSStyle.Reset
    }

    [void] run([bool]$execute) {
        $this.entries | ForEach-Object {
            $left = $_.getFullMarkerdText($true, $execute)
            $indent = $_.getLeftSideByteLen()
            $filler = $this.getFiller($indent)
            $right = $_.getFullMarkerdText($false, $execute)
            $left + $filler + $right | Write-Host
            if (-not $execute) {
                return
            }
            $item = Get-Item -LiteralPath $_.getFullname()
            $newName = $_.getNewName()
            try {
                $item | Rename-Item -NewName $newName -ErrorAction Stop
            }
            catch {
                "failed to rename as '{0}'!" -f $newName | Write-Host -ForegroundColor Magenta
            }
        }
    }
}


function Rename-ApplyScriptBlock {
    <#
        .EXAMPLE
        Get-Childitem *.txt | renB -renameBlock {$_.Name -Replace "hogehoge", "hugahuga"}
        .EXAMPLE
        Get-Childitem *.txt | Where-Object Name -Match "foo" | renB -renameBlock {$_.Name -Replace "hogehoge", "hugahuga"} -execute
    #>
    [OutputType([System.Void])]
    param (
        [scriptblock]$renameBlock
        ,[switch]$execute
    )

    $replacer = [NameReplacer]::new()
    $cur = (Get-Location).Path
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $newName = & $renameBlock
        $ent = [NameReplaceEntry]::new($_.Fullname, $cur, $newName)
        if ($ent.isRenamable()) {
            $replacer.setEntry($ent)
        }
    }
    $replacer.run($execute)
}

function Rename-LightroomFromDropbox {
    param (
        [switch]$execute
    )
    $input | Rename-ApplyScriptBlock {
        $fmt = ($_.BaseName -replace "[ \-]","" -replace "写真" -replace "\(","-" -replace "\)")
        $newName = $fmt.substring(0,8) + "-IMG_" + $fmt.substring(8).PadLeft(6, "0") + $_.extension
        return $newName
    } -execute:$execute
}

function Rename-FromData {
    [OutputType([System.Void])]
    param (
        [string[]]$data
        ,[string]$separator = "`t"
        ,[switch]$execute
    )
    $hashtable = @{}
    $data | Where-Object {$_ -replace "\s"} | ForEach-Object {
        $fields = $_ -split $separator
        $hashtable.Add($fields[0], $fields[1])
    }
    if (-not $data.Count) {
        return
    }
    $input | Where-Object {$_.Name -in $hashtable.keys} | Rename-ApplyScriptBlock { $hashtable[$_.Name] } -execute:$execute
}

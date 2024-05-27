
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
    [bool]$execute = $false
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

    [string] _getMatchesMarkeredOrgName() {
        $col = ($this.execute)? $this._execColor : "White"
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

    [string] getFullMarkerdText([bool]$org) {
        if ($org) {
            return $this._getDimmedRelDir() + $this._getMatchesMarkeredOrgName()
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
        $this._leftBufferWidth = [Math]::Max($this._leftBufferWidth, $w)
    }

    [string] _getFiller([int]$indent) {
        $rightPadding = [Math]::Max($this._leftBufferWidth - $indent, 0)
        $filler = " {0}=> " -f ("=" * $rightPadding)
        return $Global:PSStyle.Foreground.Yellow + $filler + $Global:PSStyle.Reset
    }

    [void] run() {
        $this.entries | ForEach-Object {
            $left = $_.getFullMarkerdText($true)
            $indent = $_.getLeftSideByteLen()
            $filler = $this._getFiller($indent)
            $right = $_.getFullMarkerdText($false)
            $left + $filler + $right | Write-Host
            if (-not $_.execute) {
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
    $cur = (Get-Location).ProviderPath
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [BasenameReplaceEntry]::new($_.Fullname, $cur, $from, $to, $case)
        $ent.execute = $execute
        if ($ent.isRenamable()) {
            $replacer.setEntry($ent)
        }
    }
    $replacer.run()
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
    [bool]$execute = $false
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

    [string] _getMarkerdNewName() {
        $color = ($this.execute)? "Green" : "White"
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

    [string] getFullMarkerdText() {
        return $this._getDimmedRelDir() + $this._getMarkerdNewName()
    }

}


class InsertRenamer {
    [int]$_leftBufferWidth = 0
    [InsertRenameEntry[]]$entries = @()
    InsertRenamer() {}

    [void] setEntry([InsertRenameEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        $this._leftBufferWidth = [Math]::Max($this._leftBufferWidth, $w)
    }

    [string] _getFiller([int]$leftSideLen) {
        $padding = [Math]::Max($this._leftBufferWidth - $leftSideLen, 0)
        return " " * $padding
    }

    [void] run() {
        $this.entries | ForEach-Object {
            $indent = $_.getLeftSideByteLen()
            $filler = $this._getFiller($indent)
            $filler + $_.getFullMarkerdText() | Write-Host
            if (-not $_.execute) {
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
    $cur = (Get-Location).ProviderPath
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [InsertRenameEntry]::new($_.Fullname, $cur, $insert, $position)
        $ent.execute = $execute
        if ($ent.isRenamable()) {
            $renamer.setEntry($ent)
        }
    }
    $renamer.run()
}
Set-Alias rIns Rename-Insert


class IndexRenameEntry {
    [string]$_idx
    [string]$_pre
    [string]$_indexed
    [string]$_suf
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    [bool]$hasNewName
    [bool]$execute = $false
    IndexRenameEntry([string]$fullName, [string]$curDir, [string]$altName, [int]$idx, [int]$pad, [bool]$tail) {
        $this._idx = ($idx -as [string]).PadLeft($pad, "0")
        $this._fullName = $fullName
        $item = Get-Item $this._fullName
        $this._orgBaseName = $item.BaseName
        $this._extension = $item.Extension
        $this._relDirName = [System.IO.Path]::GetRelativePath($curDir, ($this._fullName | Split-Path -Parent))
        $this.hasNewName = $altName.Length -gt 0
        if ($this.hasNewName) {
            $this._pre = ""
            $this._suf = $this._extension
            if ($tail) {
                $this._indexed = $altName + $this._idx
                return
            }
            $this._indexed = $this._idx + $altName
            return
        }
        if ($tail) {
            $this._pre = $this._orgBaseName
            $this._indexed = "_" + $this._idx
            $this._suf = $this._extension
            return
        }
        $this._pre = ""
        $this._indexed = $this._idx + "_"
        $this._suf = $this._orgBaseName + $this._extension
    }

    [string] getFullName() {
        return $this._fullName
    }

    [string] getNewName() {
        return $this._pre + $this._indexed + $this._suf
    }

    [bool] isRenamable() {
        return -not (($this._orgBaseName + $this._extension) -ceq $this.getNewName())
    }

    [string] _getMarkerdNewName() {
        $color = ($this.execute)? "Green" : "White"
        return $this._pre + `
            $Global:PSStyle.Foreground.Black + `
            $Global:PSStyle.Background.PSObject.Properties[$color].Value + `
            $this._indexed + `
            $Global:PSStyle.Reset + `
            $this._suf
    }

    [string] _getDimmedRelDir() {
        return $global:PSStyle.Foreground.BrightBlack + $this._relDirName + "\" + $global:PSStyle.Reset
    }

    [int] getPrefixByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._pre)
    }

    [int] getLeftSideByteLen() {
        $sj = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        return $sj.GetByteCount($this._relDirName + $this._orgBaseName + $this._extension)
    }

    [string] getFullMarkerdText([bool]$org) {
        if ($org) {
            return $this._getDimmedRelDir() + $this._orgBaseName + $this._extension
        }
        return $this._getDimmedRelDir() + $this._getMarkerdNewName()
    }

}

class IndexRenamer {
    [int]$_leftBufferWidth = 0
    [IndexRenameEntry[]]$entries = @()
    IndexRenamer() {}

    [void] setEntry([IndexRenameEntry]$ent) {
        $this.entries += $ent
        $w = ($ent.hasNewName)? $ent.getLeftSideByteLen() : $ent.getPrefixByteLen()
        $this._leftBufferWidth = [Math]::Max($this._leftBufferWidth, $w)
    }

    [string] _getFiller([int]$offset, [bool]$showOrigin) {
        $padding = [Math]::Max($this._leftBufferWidth - $offset, 0)
        if ($showOrigin) {
            $filler = " {0}=> " -f ("=" * $padding)
            return $Global:PSStyle.Foreground.Yellow + $filler + $Global:PSStyle.Reset
        }
        return " " * $padding
    }

    [void] run() {
        $this.entries | ForEach-Object {
            $left = ($_.hasNewName)? $_.getFullMarkerdText($true) : ""
            $offset = ($_.hasNewName)? $_.getLeftSideByteLen() : $_.getPrefixByteLen()
            $filler = $this._getFiller($offset, $_.hasNewName)
            $left + $filler + $_.getFullMarkerdText($false) | Write-Host
            if (-not $_.execute) {
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
    $cur = (Get-Location).ProviderPath
    $i = $start
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $ent = [IndexRenameEntry]::new($_.Fullname, $cur, $altName, $i, $pad, $tail)
        $ent.execute = $execute
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
    $renamer.run()
}
Set-Alias rInd Rename-Index

class NameReplaceEntry {
    [string]$_newName
    [string]$_fullName
    [string]$_orgBaseName
    [string]$_extension
    [string]$_relDirName
    [bool]$execute = $false
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

    [string] getFullMarkerdText([bool]$org) {
        if ($org) {
            return $this._getDimmedRelDir() + $this._orgBaseName + $this._extension
        }
        $color = ($this.execute)?  "Green" : "White"
        $ansi = $global:PSStyle.Foreground.PSObject.Properties[$color].Value
        return $this._getDimmedRelDir() + $ansi + $this._newName + $global:PSStyle.Reset
    }
}

class NameReplacer {
    [int]$_leftBufferWidth = 0
    [NameReplaceEntry[]]$entries = @()
    NameReplacer() {}

    [void] setEntry([NameReplaceEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getLeftSideByteLen()
        $this._leftBufferWidth = [Math]::Max($this._leftBufferWidth, $w)
    }

    [string] _getFiller([int]$indent, [bool]$execute) {
        $color = ($execute)? "Cyan" : "Black"
        $ansi = $global:PSStyle.Foreground.PSObject.Properties[$color].Value
        $rightPadding = [Math]::Max($this._leftBufferWidth - $indent, 0)
        $filler = " {0}=> " -f ("=" * $rightPadding)
        return $ansi + $filler + $Global:PSStyle.Reset
    }

    [void] run() {
        $this.entries | ForEach-Object {
            $left = $_.getFullMarkerdText($true)
            $indent = $_.getLeftSideByteLen()
            $filler = $this._getFiller($indent, $ent.execute)
            $right = $_.getFullMarkerdText($false)
            $left + $filler + $right | Write-Host
            if (-not $ent.execute) {
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
    $cur = (Get-Location).ProviderPath
    $input | Where-Object {Test-Path $_} | ForEach-Object {Get-Item $_} | ForEach-Object {
        $newName = & $renameBlock
        $ent = [NameReplaceEntry]::new($_.Fullname, $cur, $newName)
        $ent.execute = $execute
        if ($ent.isRenamable()) {
            $replacer.setEntry($ent)
        }
    }
    $replacer.run()
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

function Rename-ByRule {
    [OutputType([System.Void])]
    param (
        [string[]]$rule
        ,[string]$separator = "`t"
        ,[switch]$execute
    )
    $table = @{}
    $rule | Where-Object {$_.Trim().Length -gt 0} | ForEach-Object {
        $fields = $_ -split $separator
        $table.Add($fields[0], $fields[1])
    }
    if (-not $table.Count) {
        return
    }
    $input | Where-Object {$_.Name -in $table.keys} | Rename-ApplyScriptBlock { $table[$_.Name] } -execute:$execute
}

class PSStringDiff {
    [string]$_decoReset = $global:PSStyle.Reset
    [string]$_decoAdd = $global:PSStyle.Background.BrightBlue + $global:PSStyle.Foreground.Black
    [string]$_decoDel = $global:PSStyle.Background.BrightRed + $global:PSStyle.Strikethrough + $global:PSStyle.Foreground.Black
    [PSObject[]]$_deltas
    PSStringDiff([string]$fromStr, [string]$toStr) {
        $froms = $fromStr.GetEnumerator().ForEach({$_ -as [string]})
        $tos = $toStr.GetEnumerator().ForEach({$_ -as [string]})
        $this._deltas = Compare-Object -ReferenceObject $froms -DifferenceObject $tos -IncludeEqual -CaseSensitive -SyncWindow 0
    }

    [string] _markup([int]$idx) {
        $d = $this._deltas[$idx]
        $t = $d.InputObject
        if ($d.SideIndicator -eq "=>") {
            return $this._decoAdd + $t + $this._decoReset
        }
        if ($d.SideIndicator -eq "<=") {
            return $this._decoDel + $t + $this._decoReset
        }
        return $t
    }

    [string] execute() {
        $builder = [System.Text.StringBuilder]::new()
        for ($i = 0; $i -lt $this._deltas.Count; $i++) {
            $s = $this._markup($i)
            $builder.Append($s) > $null
        }
        return $builder.ToString()
    }
}

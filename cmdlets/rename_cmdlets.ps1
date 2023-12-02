
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

    [int] getIndentDepth() {
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
    [int]$bufferWidth = 0
    [BasenameReplaceEntry[]]$entries = @()
    BasenameReplaceDiffer() {}

    [void] setEntry([BasenameReplaceEntry]$ent) {
        $this.entries += $ent
        $w = $ent.getIndentDepth()
        if ($this.bufferWidth -lt $w) {
            $this.bufferWidth = $w
        }
    }

    [string] getFiller([int]$indent) {
        $rightPadding = $this.bufferWidth - $indent
        if ($rightPadding -lt 0) {
            $rightPadding = 0
        }
        $filler = " {0}=> " -f ("=" * $rightPadding)
        return $Global:PSStyle.Foreground.Yellow + $filler + $Global:PSStyle.Reset
    }

    run([bool]$execute) {
        $this.entries | ForEach-Object {
            $left = $_.getFullMarkerdText($true, $execute)
            $indent = $_.getIndentDepth()
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
    $input | Where-Object {Test-Path $_} | ForEach-Object {
        $ent = [BasenameReplaceEntry]::new($_.Fullname, $cur, $from, $to, $case)
        if ($ent.isRenamable()) {
            $replacer.setEntry($ent)
        }
    }
    $replacer.run($execute)
}
Set-Alias repBN Rename-ReplaceBasename



Class RenamePreview {
    [int]$bufferWidth
    [System.Text.Encoding]$sjis

    RenamePreview($targets) {
        $this.sjis = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        $this.bufferWidth = $targets | ForEach-Object {$this.sjis.GetByteCount($_)} | Sort-Object | Select-Object -Last 1
    }

    [string] Fill([string]$subStr, [bool]$arrow) {
        $trimLen = 0
        [regex]::new("`e\[.+?m").Matches($subStr).ForEach({ $trimLen += $_.Length })
        $fillerWidth = $this.bufferWidth - ($this.sjis.GetByteCount($subStr) - $trimLen)
        if ($fillerWidth -lt 0) {
            $fillerWidth = 0
        }
        $filler = ($arrow)? (" {0}=> " -f ("=" * $fillerWidth)) : (" " * $fillerWidth)
        return $Global:PSStyle.Foreground.BrightBlack + $filler + $Global:PSStyle.Reset
    }

    [string] Compare([string]$before, [string]$after, [string]$color, [bool]$arrow) {
        return $before + $this.Fill($before, $arrow) + $Global:PSStyle.Foreground.PSObject.Properties[$color].Value + $after + $Global:PSStyle.Reset
    }

}


class InsertRenamer {
    [string]$pre
    [string]$post
    [string]$extension
    [string]$insert

    InsertRenamer([string]$path, [int]$pos, [string]$insert) {
        $file = Get-Item -LiteralPath $path
        $this.extension = $file.Extension
        if ([Math]::Abs($pos) -gt $file.BaseName.Length + 1) {
            $this.pre = $file.Basename
            $this.post = ""
            $this.insert = ""
            return
        }
        $this.insert = $insert
        if ($pos -eq 0) {
            $this.pre = ""
            $this.post = $file.Basename
            return
        }
        if ($pos -gt 0) {
            $this.pre = ($file.BaseName).substring(0, $pos)
            $this.post = ($file.BaseName).substring($pos)
            return
        }
        $this.pre = ($file.BaseName).substring(0, ($file.BaseName).length + $pos + 1)
        $this.post = ($file.BaseName).substring(   ($file.BaseName).length + $pos + 1)
    }

    [string] GetText() {
        return $this.pre + $this.insert + $this.post + $this.extension
    }

    [string] GetMarkup([string]$color) {
        return $this.pre + $Global:PSStyle.Foreground.Black + $Global:PSStyle.Background.PSObject.Properties[$color].Value + $this.insert + $Global:PSStyle.Reset + $this.post + $this.extension
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

    $procTarget = $input | Where-Object {$_.GetType().Name -in @("FileInfo", "DirectoryInfo")}
    if (-not $procTarget) {
        return
    }
    $previewer = [RenamePreview]::new($procTarget.Name)
    $color = ($execute)? "Green" : "White"

    $procTarget | ForEach-Object {
        $rn = [InsertRenamer]::new($_.fullname, $position, $insert)
        if ($position -lt 0) {
            $previewer.Fill($_.Name, $false) | Write-Host -NoNewline
        }
        $rn.GetMarkup($color) | Write-Host
        if ($execute) {
            $_ | Rename-Item -NewName $rn.GetText()
        }
    }
}
Set-Alias rIns Rename-Insert

class IndexedItem {
    [string]$pre
    [string]$post
    [string]$modified

    IndexedItem([string]$path, [int]$i, [int]$padding, [string]$altName, [bool]$tail) {
        $idx = ($i -as [string]).PadLeft($padding, "0")
        $file = Get-Item -LiteralPath $path
        if ($altName) {
            $this.pre = ""
            $this.post = $file.Extension
            if ($tail) {
                $this.modified = $altName + $idx
                return
            }
            $this.modified = $idx + $altName
            return
        }
        if ($tail) {
            $this.pre = $file.BaseName
            $this.modified = "_" + $idx
            $this.post = $file.Extension
            return
        }
        $this.pre = ""
        $this.modified = $idx + "_"
        $this.post = $file.Name
    }

    [string] GetText() {
        return $this.pre + $this.modified + $this.post
    }

    [string] GetMarkup($color) {
        return $this.pre + $Global:PSStyle.Foreground.Black + $Global:PSStyle.Background.PSObject.Properties[$color].Value + $this.modified + $Global:PSStyle.Reset + $this.post;
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
        ,[int[]]$skip
        ,[switch]$tail
        ,[switch]$execute
    )

    $proc = $input | Where-Object {$_.GetType().Name -in @("FileInfo", "DirectoryInfo")}
    $previewer = [RenamePreview]::new($proc.Name)
    $color = ($execute)? "Green" : "White"

    $idx = $start - $step
    $proc | ForEach-Object {
        $idx += $step
        if ($skip.Length) {
            while ($idx -in $skip) {
                $idx += $step
            }
        }

        $ixi = [IndexedItem]::new($_.Fullname, $idx, $pad, $altName, $tail)
        $markup = $ixi.GetMarkup($color)

        if ($altName) {
            $previewer.Compare($_.Name, $markup, $color, $true) | Write-Host
        }
        else {
            if ($tail) {
                $previewer.Fill($_.Name, $false) | Write-Host -NoNewline
            }
            $markup | Write-Host
        }

        if ($execute) {
            $_ | Rename-Item -NewName $ixi.GetText()
        }

    }
}
Set-Alias rInd Rename-Index

function Rename-WithScriptBlock {
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
    $procTarget = $input | Where-Object {$_.GetType().Name -in @("FileInfo", "DirectoryInfo")}

    $previewer = [RenamePreview]::new($procTarget.Name)
    $color = ($execute)? "Green" : "White"

    $procTarget | ForEach-Object {
        $newName = & $renameBlock
        $previewer.Compare($_.Name, $newName, $color, $true) | Write-Host
        if ($execute) {
            $_ | Rename-Item -NewName $newName
        }
    }
}

function Rename-LightroomFromDropbox {
    param (
        [switch]$execute
    )
    $input | Rename-WithScriptBlock {
        $fmt = ($_.BaseName -replace "[ \-]","" -replace "写真" -replace "\(","-" -replace "\)")
        $newName = $fmt.substring(0,8) + "-IMG_" + $fmt.substring(8).PadLeft(6, "0") + $_.extension
        return $newName
    } -execute:$execute
}


function Rename-Format {
    <#
        .EXAMPLE
        Get-Childitem *.txt | Rename-Format "01_{0}_test"
    #>
    [OutputType([System.Void])]
    param (
        [string]$format = "{0}"
        ,[switch]$execute
    )
    $procTarget = $input | Where-Object {$_.GetType().Name -in @("FileInfo", "DirectoryInfo")}
    $previewer = [RenamePreview]::new($procTarget.Name)
    $color = ($execute)? "Green" : "White"

    $procTarget | ForEach-Object {
        $newName = ($format -f $_.Basename) + $_.Extension
        $previewer.Compare($_.Name, $newName, $color, $true) | Write-Host
        if ($execute) {
            $_ | Rename-Item -NewName $newName
        }
    }
}
Set-Alias renFmt Rename-Format

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
    $input | Where-Object {$_.Name -in $hashtable.keys} | Rename-WithScriptBlock { $hashtable[$_.Name] } -execute:$execute
}

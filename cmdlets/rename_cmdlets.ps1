
<# ==============================

cmdlets for renaming file or folder

                encoding: utf8bom
============================== #>

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

    $reg = ($case)? [regex]::new($from) : [regex]::new($from, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    $procTarget = $input | Where-Object {$reg.IsMatch($_.Basename)}
    if ($procTarget.Count -lt 1) {
        return
    }

    $color = ($execute)? "Cyan" : "White"
    $hi = [PsHighlight]::new($from, $color, $case)
    $previewer = [RenamePreview]::new($procTarget.Name)

    $procTarget | ForEach-Object {
        $newName = $reg.Replace($_.Basename, $to) + $_.Extension
        $markup = $hi.Markup($_.Basename) + $_.Extension
        $previewer.Compare($markup, $newName, $color, $true) | Write-Host
        if ($execute) {
            try {
                $_ | Rename-Item -NewName $newName -ErrorAction Stop
            }
            catch {
                "same file '{0}' already exists!" -f $newName | Write-Host -ForegroundColor Magenta
            }
        }
    }
}
Set-Alias repBN Rename-ReplaceBasename

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

class IndexRenamer {
    [string]$pre
    [string]$post
    [string]$modified

    IndexRenamer([string]$path, [int]$i, [int]$padding, [string]$altName, [bool]$tail) {
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
        while ($skip.Length -and $idx -in $skip) {
            $idx += $step
        }

        $ir = [IndexRenamer]::new($_.Fullname, $idx, $pad, $altName, $tail)
        $markup = $ir.GetMarkup($color)

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
            $_ | Rename-Item -NewName $ir.GetText()
        }

    }
}
Set-Alias rInd Rename-Index

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
Set-Alias renAS Rename-ApplyScriptBlock


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
    $input | Where-Object {$_.Name -in $hashtable.keys} | Rename-ApplyScriptBlock { $hashtable[$_.Name] } -execute:$execute
}


<# ==============================

cmdlets for formating text

                encoding: utf8bom
============================== #>

function ConvertTo-HexString {
    $input | ForEach-Object {
        if ($_ -as [int]) {
            return [string]("{0:X}" -f ($_ -as [int]))
        }
        return [string]$_
    }
}

function Join-StringUtil {
    param (
        [string]$connector = ""
        ,[string]$wrapper = ""
        ,[string]$suffix = ""
        ,[string]$prefix = ""
    )
    if($wrapper.Length -eq 2) {
        $pre = $wrapper[0]
        $post = $wrapper[1]
    }
    else {
        $pre = $post = $wrapper
    }
    $input | ForEach-Object {$prefix + $_ + $suffix} | Join-String -Separator "$post$connector$pre" -OutputPrefix $pre -OutputSuffix $post | Write-Output
}
Set-Alias jnt Join-StringUtil

function Format-String {
    param (
        [string]$format = "{0}"
        )
        $input | ForEach-Object {
            $format -f $_ | Write-Output
        }
    }
Set-Alias fmt Format-String


function Format-ReplaceLine {
    param (
        [string]$from
        ,[string]$to = ""
        ,[switch]$case
        ,[switch]$simple
    )
    $lines = New-Object System.Collections.ArrayList
    @($input).ForEach({$lines.Add($_) > $null})

    $lines | ForEach-Object {
        $line = $_
        $replaced = & {
            if ($simple) {
                $opt = ($case)? [System.StringComparison]::Ordinal : [System.StringComparison]::OrdinalIgnoreCase
                return ($line -as [string]).Replace($from, $to, $opt)
            }
            if ($case) {
                return $line -creplace $from, $to
            }
            return $line -replace $from, $to
        }
        [System.Console]::ForegroundColor = $(($line -ceq $replaced)? "Gray" : "Blue")
        Write-Output $replaced
        [System.Console]::ResetColor()
    }
}
Set-Alias rl Format-ReplaceLine

function Format-MatchLine {
    param (
        [scriptblock]$matchBlock = {return $true}
        ,[scriptblock]$formatBlock = {$_}
        ,[scriptblock]$elseBlock = {$_}
    )
    $lines = New-Object System.Collections.ArrayList
    @($input).ForEach({$lines.Add($_) > $null})

    $lines | ForEach-Object {
        if ($matchBlock.Invoke()){
            $formatBlock.Invoke() | Write-Output
        }
        else {
            $elseBlock.Invoke() | Write-Output
        }
    }
}
Set-Alias fmtMch Format-MatchLine


function ConvertTo-SequenceByValue {
    $arrayList = New-Object System.Collections.ArrayList
    @($input).ForEach({$arrayList.Add($_) > $null})
    $counter = 1
    for ($i = 0; $i -lt $arrayList.Count; $i++) {
        if ($i -eq 0) {
            $counter | Write-Output
            continue
        }
        $pre = $arrayList[$i - 1]
        $cur = $arrayList[$i]
        if ($pre -cne $cur) {
            $counter += 1
        }
        $counter | Write-Output
    }
}

function ConvertTo-IncrementalSequence {
    param (
        [int]$start = 1
    )
    $input |ForEach-Object {
        $start | Write-Output
        $start += 1
    }
}
Set-Alias toSeq ConvertTo-IncrementalSequence

function Format-InsertIndex {
    param (
        [int]$position = 0
        ,[int]$start = 1
        ,[int]$pad = 1
    )
    $idx = $start - 1
    $input | ForEach-Object {
        $idx += 1
        $maxLen = $_.Length
        if ([Math]::Abs($position) -gt $maxLen) {
            return $_
        }
        if ($position -ge 0) {
            $prefix = $_.Substring(0, $position)
            $suffix = $_.Substring($position, ($maxLen - $position))
        }
        else {
            $pos = $position + 1
            $prefix = $_.Substring(0, $maxLen + $pos)
            $suffix = $_.Substring(($maxLen + $pos), [Math]::Abs($pos))
        }
        return $prefix + ($idx -as [string]).PadLeft($pad, "0") + $suffix
    }

}

function Format-ReplaceNth {
    <#
        .EXAMPLE
        cat \.hoge.txt | Format-ReplaceNth -n 3 -to "aaa"
    #>
    param (
        [int]$n,
        [string]$to
    )

    if ($n -eq 0) {
        return
    }
    else {
        $len = [math]::Abs($n) - 1
    }

    $reg = ($n -gt 0)? [regex]"(?<=^.{$len}).": [regex]".(?=.{$len}$)"
    $input | ForEach-Object {
        $reg.Replace($_, $to)
    }
}
Set-Alias replN Format-ReplaceNth

function Select-MatchLine {
    <#
        .EXAMPLE
        $hoge | slm "fuga" # equal to: $hoge | sls "fuga" | foreach line
        $hoge | slm "piyo" -selectNotMatch # equal to: $hoge | where $_ -notmatch "piyo"
    #>
    param (
        [string]$pattern
        ,[switch]$case
        ,[switch]$selectNotMatch
        ,[int]$stepFromMatch = 0
    )

    if ($selectNotMatch) {
        return $($input | Select-String -Pattern $pattern -CaseSensitive:$case -NoEmphasis -NotMatch)
    }
    $abs = [math]::Abs($stepFromMatch) -as [int]
    $grep = $input | Select-String -Pattern $pattern -CaseSensitive:$case -NoEmphasis -Context $abs
    if ($stepFromMatch -gt 0) {
        return $($grep | ForEach-Object {($_.Context.PostContext)[($stepFromMatch - 1)]})
    }
    elseif ($stepFromMatch -lt 0) {
        return $($grep | ForEach-Object {($_.Context.PreContext)[($abs + $stepFromMatch)]})
    }
    return $grep
}
Set-Alias slm Select-MatchLine

function Convert-PSObjectToTable {
    <#
        .EXAMPLE
        (gcb) | fromTSV -header aaa,bbb,ccc | Convert-PSObjectToTable
    #>
    $markup = $input | ConvertTo-Html
    Write-Output "<table>"
    Write-Output "<thead>"
    $markup | Select-Object -Index 7 | Write-Output
    Write-Output "</thead>"
    $markup | Select-Object -Skip 8 | Where-Object {$_ -match "^<tr>"} | ForEach-Object { $_ -replace "\r?\n" , "<br />" } | Write-Output
    Write-Output "</table>"
}


function Convert-Tsv2MarkdownTable {
    <#
        .EXAMPLE
        (gcb) | Convert-Tsv2MarkdownTable
    #>
    param (
        [ValidateSet("left", "center", "right")][string]$align = "left"
    )
    $indicator = @{
        "left" = ":---";
        "center" = ":---:";
        "right" = "---:";
    }[$align]
    $tsv = $input | ConvertFrom-Csv -Delimiter "`t"
    $ths = ($tsv | Select-Object -First 1).PsObject.Properties.Name
    $ths | Join-String -Separator "|" -OutputPrefix "|" -OutputSuffix "|" | Write-Output
    $ths | ForEach-Object {$indicator} | Join-String -Separator "|" -OutputPrefix "|" -OutputSuffix "|" | Write-Output
    $tsv | ForEach-Object {
        $_.PsObject.Properties.Value -replace "\r?\n", "<br>" | ForEach-Object {
            return ($_)? $_ : " "
        } | Join-String -Separator "|" -OutputPrefix "|" -OutputSuffix "|" | Write-Output
    }
}
Set-Alias mdTableFromTSV Convert-Tsv2MarkdownTable


function Format-StripPostalCodeFromAddress {
    <#
        .EXAMPLE
        gcb | Format-StripPostalCodeFromAddress
    #>
    $reg = [regex]"〒?(\d{3}).(\d{4})\s*(.+$)"
    $input | ForEach-Object {
        $s = ($_ -as [string]).Trim()
        $m = $reg.Match($s)
        if ($m.Success) {
            $groups = $m.Groups
            $postalcode = "{0}-{1}" -f $groups[1].Value, $groups[2].Value
            $address = $groups[3]
            return [PSCustomObject]@{
                "Postalcode" = $postalcode;
                "Adddress" = $address;
                "TSV" = $postalcode + "`t" + $address;
            }
        }
        return [PSCustomObject]@{
            "Postalcode" = "";
            "Adddress" = $_;
            "TSV" = $_;
        }
    }
}

function ConvertTo-IdNumber {
    <#
        .SYNOPSIS
        パイプライン経由で入力される要素をもとに、値が変わるたびに1ずつ増える数字の列を生成する
        ・Microsoft Excel でストライプ模様を作成する際などに使用。
    #>
    param (
        [int]$start = 1
    )
    $inputArray = New-Object System.Collections.ArrayList
    $input.ForEach({
        $inputArray.Add($_) > $null
    })
    [int]$index = $start
    Write-Output $index
    for ($i = 1; $i -lt $inputArray.Count; $i++) {
        $current = $inputArray[$i]
        $previous = $inputArray[($i - 1)]
        if ($current -ne $previous) {
            $index += 1
        }
        Write-Output $index
    }
}

function Format-EmbedYoutube {
    param (
        [parameter(ValueFromPipeline)]$inputLine
    )
    begin {}
    process {
        $inputLine -replace "^.+\.youtube\.com/watch\?v=(.{11}).*", '<div class="youtube"><iframe src="https://www.youtube.com/embed/$1?mute=1&rel=0" frameborder="0" allowfullscreen></iframe></div>' | Write-Output
    }
    end {}
}

function Get-SortInfo {
    param (
        [parameter(ValueFromPipeline)][string]$inputLine
    )
    <#
        .EXAMPLE
        Get-SortInfo "Akita, W., & Tanaka, I. (2000)."
        => Akita_Tanaka_2000
    #>
    begin {}
    process {
        $arr = @($inputLine -split "\d{4}")
        if ($arr.Count -lt 2) {
            return $inputLine
        }
        $names, $rest = $arr
        $y = $inputLine.Substring($names.Length).Substring(0, 4)
        $fmt = ""
        if ($names -match "^[a-z]") {
            $fmt = $names -replace "[^\s]+\." -replace "[&,]"
        }
        else {
            $fmt = $names -replace "[・／]"
        }
        return "$($fmt -replace "\s+", "_")_$y" -replace "[（\(]_", "_"
    }
    end {}
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "ToRtfHighlight" -Value {
    # https://latex2rtf.sourceforge.net/rtfspec_6.html
    # https://www.biblioscape.com/rtf15_spec.htm#Heading45
    $colortbl = "{\colortbl;"
    @(
        @(0, 0, 0),
        @(0, 0, 255),
        @(0, 255, 255),
        @(0, 255, 0),
        @(255, 0, 255),
        @(255, 0, 0),
        @(255, 255, 0),
        @(255, 255, 255),
        @(0, 0, 128),
        @(0, 128, 128),
        @(0, 128, 0),
        @(128, 0, 128),
        @(128, 0, 0),
        @(128, 128, 0),
        @(128, 128, 128),
        @(192, 192, 192)
    ) | ForEach-Object {
        $colortbl += ("\red{0}\green{1}\blue{2};" -f $_[0], $_[1], $_[2])
    }
    $colortbl += "}"
    $t = "{\i\cf1\highlight7 " + $this + "}"
    return $colortbl + $t
}

function ConvertTo-Rtf {
    param (
        [parameter(ValueFromPipeline = $true)]$inputLine
    )
    begin {}
    process {
        "{\rtf" + $inputLine + "}" | Write-Output
    }
    end {}
}

function Set-ClipboardAsRtf {
    <#
    .EXAMPLE
        "aa" + "bb".ToRtfHighlight() + "cc" | ConvertTo-Rtf | Set-ClipboardAsRtf
    #>
    param (
        [parameter(ValueFromPipeline = $true)]$inputLine
    )
    begin {
        $lines = @()
    }
    process {
        $lines += $inputLine
    }
    end {
        $rtf = $lines -join "`n"
        [System.Windows.Forms.Clipboard]::SetText($rtf, [System.Windows.Forms.TextDataFormat]::Rtf)
    }
}

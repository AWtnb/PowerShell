
<# ==============================

cmdlets for processing japanese

            encoding: utf8bom
============================== #>

function googleDict {
    $env:USERPROFILE | Join-Path -ChildPath "Sync\develop\app_setting\IME_google\convertion_dict\main.txt" | Get-Item | Get-Content | bat.exe -p
}

function aw {
    "あかさたなはまやらわ".GetEnumerator() | Write-Output
}


####################
# Class: FormatJP
####################

class FormatJP {

    static [string] ToHiragana ([string]$strData) {
        $converted = $strData
        ([regex]"[\u30a1-\u30f6]").Matches($strData).Value | Where-Object {$_} | Sort-Object | Get-Unique | ForEach-Object {
            $c = [int]($_ -as [char]) - 96
            $hira = [char]::ConvertFromUtf32($c)
            $converted = $converted -replace $_, $hira
        }
        return $converted
    }

    static [string] ToKatakana ([string]$strData) {
        $converted = $strData
        ([regex]"[\u3041-\u3096]").Matches($strData).Value | Where-Object {$_} | Sort-Object | Get-Unique | ForEach-Object {
            $c = [int]($_ -as [char]) + 96
            $kata = [char]::ConvertFromUtf32($c)
            $converted = $converted -replace $_, $kata
        }
        return $converted
    }

    static [string] ToVoicing ([string]$strData) {
        $converted = $strData
        [regex]::Matches($converted, "[カキクケコサシスセソタチツテトハヒフヘホかきくけこさしすせそたちつてとはひふへほ]").Value | Where-Object {$_} | ForEach-Object {
            $converted = $converted -replace $_, [string]([Convert]::ToChar([Convert]::ToInt32([char]$_) + 1))
        }
        return $($converted -replace "う", "ゔ" -replace "ウ", "ヴ")
    }

    static [string] ToHalfVoicing ([string]$strData) {
        $converted = $strData
        [regex]::Matches($converted, "[ハヒフヘホはひふへほ]").Value | Where-Object {$_} | ForEach-Object {
            $converted = $converted -replace $_, [string]([Convert]::ToChar([Convert]::ToInt32([char]$_) + 2))
        }
        return $converted
    }

    static [string] ToVoiceless ([string]$strData) {
        $converted = $strData
        [regex]::Matches($converted, "[ガギグゲゴザジズゼゾダヂヅデドバビブベボがぎぐげこざじずぜぞだぢづでどばびぶべぼ]").Value | Where-Object {$_} | ForEach-Object {
            $withoutVoicing = [Convert]::ToChar([Convert]::ToInt32([char]$_) - 1) -as [string]
            $converted = $converted -replace $_, $withoutVoicing
        }
        [regex]::Matches($converted, "[パピプペポぱぷぷぺぽ]").Value | Where-Object {$_} | ForEach-Object {
            $withoutHalfVoicing = [Convert]::ToChar([Convert]::ToInt32([char]$_) - 2) -as [string]
            $converted = $converted -replace $_, $withoutHalfVoicing
        }
        return $($converted -replace "\u30f4", "ウ" -replace "\u3094", "う")
    }

    static [string] Normalize ([string]$strData) {
        $dict = @{
            "ァ" = "ア";
            "ィ" = "イ";
            "ゥ" = "ウ";
            "ェ" = "エ";
            "ォ" = "オ";
            "ッ" = "ツ";
            "ャ" = "ヤ";
            "ュ" = "ユ";
            "ョ" = "ヨ";
            "ー" = "";
        }
        $converted = [FormatJP]::ToVoiceless([FormatJP]::ToKatakana($strData)) -replace "\W"
        foreach ($key in $dict.Keys) {
            $converted = $converted -replace $key, $dict[$key]
        }
        return $converted
    }

    static [string] ToRoman ([string]$strData) {
        $dict = @{
            "ア"="A";"イ"="I";"ウ"="U";"エ"="E";"オ"="O";
            "カ"="Ka";"キ"="Ki";"ク"="Ku";"ケ"="Ke";"コ"="Ko";
            "サ"="Sa";"シ"="Shi";"ス"="Su";"セ"="Se";"ソ"="So";
            "タ"="Ta";"チ"="Chi";"ツ"="Tsu";"テ"="Te";"ト"="To";
            "ナ"="Na";"ニ"="Ni";"ヌ"="Nu";"ネ"="Ne";"ノ"="No";
            "ハ"="Ha";"ヒ"="Hi";"フ"="Fu";"ヘ"="He";"ホ"="Ho";
            "マ"="Ma";"ミ"="Mi";"ム"="Mu";"メ"="Me";"モ"="Mo";
            "ヤ"="Ya";"ユ"="Yu";"ヨ"="Yo";
            "ラ"="Ra";"リ"="Ri";"ル"="Ru";"レ"="Re";"ロ"="Ro";
            "ワ"="Wa";"ヲ"="Wo";"ン"="N";
            "ガ"="Ga";"ギ"="Gi";"グ"="Gu";"ゲ"="Ge";"ゴ"="Go";
            "ザ"="Za";"ジ"="Ji";"ズ"="Zu";"ゼ"="Ze";"ゾ"="Zo";
            "ダ"="Da";"ヂ"="Di";"ヅ"="Zu";"デ"="De";"ド"="Do";
            "バ"="Ba";"ビ"="Bi";"ブ"="Bu";"ベ"="Be";"ボ"="Bo";
            "パ"="Pa";"ピ"="Pi";"プ"="Pu";"ペ"="Pe";"ポ"="Po";
            "ャ"="Lya";"ュ"="Lyu";"ョ"="Lyo";"ッ"="Ltu";
        }
        $converted = [FormatJP]::ToKatakana($strData)
        foreach ($key in $dict.Keys) {
            $converted = $converted -replace $key, $dict[$key]
        }
        # サ行タ行の拗音処理 → 拗音処理 → 促音処理
        $converted = $converted -replace "([CS]h|J)iLy(.)", '$1$2' -replace "([A-Z])iL(y.)", '$1$2' -replace "Ltu(.)", '$1$1'
        return $converted.ToLower()
    }

    static [string] ToHalfWidth ([string]$strData) {
        $converted = $strData
        [regex]::Matches($converted, "[ａ-ｚＡ-Ｚ０-９]").Value | Where-Object {$_} | Sort-Object | Get-Unique | ForEach-Object {
            $c = [int]($_ -as [char]) - 65248
            $halfWidth = [char]::ConvertFromUtf32($c)
            $converted = $converted -replace $_, $halfWidth
        }
        return $($converted -replace "（","(" -replace "）",")")
    }

    static [string] ToFullWidth ([string]$strData) {
        $converted = $strData
        [regex]::Matches($converted, "[a-zA-Z0-9]").Value | Where-Object {$_} | Sort-Object | Get-Unique | ForEach-Object {
            $c = 65248 + [int]($_ -as [char])
            $fullWidth = [char]::ConvertFromUtf32($c)
            $converted = $converted -replace $_, $fullWidth
        }
        return $($converted -replace "\(","（" -replace "\)","）")
    }

    static [string] FromRoman ([string]$strData) {
        $converted = $strData
        @{
            "A" = "えい";
            "B" = "ひ";
            "C" = "し";
            "D" = "てい";
            "E" = "い";
            "F" = "えふ";
            "G" = "し";
            "H" = "えいち";
            "I" = "あい";
            "J" = "しえい";
            "K" = "けい";
            "L" = "える";
            "M" = "えむ";
            "N" = "えぬ";
            "O" = "お";
            "P" = "ひ";
            "Q" = "きゆ";
            "R" = "ある";
            "S" = "えす";
            "T" = "てい";
            "U" = "ゆ";
            "V" = "ふい";
            "W" = "たふりゆ";
            "X" = "えくす";
            "Y" = "わい";
            "Z" = "せつと";
        }.GetEnumerator() | ForEach-Object {
            $converted = $converted -replace $_.key, $_.value
        }
        return $converted
    }

}

[FormatJP].DeclaredMembers | Where-Object MemberType -eq Method | ForEach-Object {
    @"
Update-TypeData -MemberName $($_.Name) -TypeName System.String -Force -MemberType ScriptMethod -Value {
    return [FormatJP]::$($_.Name)(`$this)
}
"@ | Invoke-Expression
}

@{
    "ConvertTo-Katakana"  = "ToKatakana";
    "ConvertTo-Hiragana"  = "ToHiragana";
    "ConvertTo-Hairetsu"  = "Normalize";
    "ConvertFrom-Roman"   = "FromRoman";
    "ConvertTo-Roman"     = "ToRoman";
    "ConvertTo-HalfWidth" = "ToHalfWidth";
    "ConvertTo-FullWidth" = "ToFullWidth";
}.GetEnumerator() | ForEach-Object { @"
function $($_.Key) {
    param (
        [parameter(ValueFromPipeline = `$true)]`$s
    )
    begin {}
    process {
        [FormatJP]::$($_.Value)(`$s) | Write-Output
    }
    end {}
}
"@ | Invoke-Expression }

Update-TypeData -MemberName "CountMatches" -TypeName System.String -Force -MemberType ScriptMethod -Value {
    param ($pattern=".", $case=$true)
    $opt = ($case)? [System.Text.RegularExpressions.RegexOptions]::None : [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    return [regex]::Matches($this, $pattern, $opt).Count
}

####################
# Class: Unicode
####################

class Unicode {

    static [string] ToLetter ([string]$strData) {
        return [string](
            [char]::ConvertFromUtf32(
                [Convert]::ToInt32($strData, 16)
            )
        )
    }

    static [string] FromLetter ([string]$strData) {
        return [string](
            [Convert]::ToInt32($strData -as [char]).ToString("x")
        )
    }

}

Update-TypeData -MemberName ToUnicode -TypeName System.String -Force -MemberType ScriptMethod -Value {
    return $this.GetEnumerator() | ForEach-Object { [Unicode]::FromLetter($_) }
}
Update-TypeData -MemberName FromUnicode -TypeName System.String -Force -MemberType ScriptMethod -Value {
    return [Unicode]::ToLetter($this)
}

function Convert-UnicodeToLetter {
    param (
        [parameter(ValueFromPipeline)]$code
    )
    begin {}
    process {
        [Unicode]::ToLetter($code) | Write-Output
    }
    end {}
}

function Convert-LetterToUnicode {
    param (
        [parameter(ValueFromPipeline)]$s
    )
    begin {}
    process {
        $s.GetEnumerator() | ForEach-Object {
            [Unicode]::FromLetter($_) | Write-Output
        }
    }
    end {}
}

class SudachiTokenWrapper {
    [string]$reading = ""
    [string]$detail = ""

    SudachiTokenWrapper($token) {
        $surface = $token.surface
        if ($token.pos -match "記号" -or $token.pos -match "空白" -or $surface -match "^([ぁ-んァ-ヴ・ー]|[a-zA-Zａ-ｚＡ-Ｚ]|[0-9０-９]|[\W\s])+$") {
            if ($surface -match "[ぁ-ん]") {
                $this.reading = [FormatJP]::ToKatakana($surface)
            }
            else {
                $this.reading = $surface
            }
            $this.detail = $surface
            return
        }
        if (-not $token.reading) {
            $this.reading = $surface
            $this.detail = "{0}(?)" -f $surface
            return
        }
        $this.reading = $token.reading
        $this.detail = "{0}({1})" -f $surface, $this.reading
    }

}

class SudachiTokensReader {
    [PSCustomObject[]]$tokens

    SudachiTokensReader($rawTokens) {
        $this.tokens = $rawTokens.ForEach({[SudachiTokenWrapper]::new($_)})
    }

    [string] GetReading() {
        $builder = New-Object System.Text.StringBuilder
        foreach ($token in $this.tokens) {
            $builder.Append($token.reading) > $null
        }
        return $builder.ToString()
    }

    [string] GetDetail() {
        $stack = New-Object System.Collections.ArrayList
        foreach ($token in $this.tokens) {
            $stack.Add($token.detail) > $null
        }
        return $stack -join " / "
    }


}

class SudachiPy {
    [PSCustomObject[]]$parsed

    SudachiPy([string[]]$lines, [bool]$ignoreParen=$false, [bool]$focusName=$false) {
        $sudachiPath = $PSScriptRoot | Join-Path -ChildPath "python\sudachi_tokenizer.py"
        Use-TempDir {
            $in = New-Item -Path ".\in.txt"
            $out = New-Item -Path ".\out.txt"
            $lines | Out-File -Encoding utf8NoBOM -FilePath $in.FullName
            $opt = @()
            $opt += (($ignoreParen)? "IgnoreParen" : "IncludeParen")
            $opt += (($focusName)? "FocusName" : "IncludeNoise")
            Start-Process -Path python.exe -wait -ArgumentList (@("-B", $sudachiPath, $in.FullName, $out.FullName) + $opt) -NoNewWindow
            $this.parsed = Get-Content -Path $out.FullName | ConvertFrom-Json
        }
    }

    [PSCustomObject[]] GetReading() {
        return $this.parsed | ForEach-Object {
            $reader = [SudachiTokensReader]::new($_.tokens)
            $line = $_.raw_line
            $reading = $reader.GetReading()
            return [PSCustomObject]@{
                "Line"       = $line;
                "Reading"    = $reading;
                "Tokenize"   = $reader.GetDetail();
                "Normalized" = [FormatJP]::Normalize($reading);
                "Roman"      = [FormatJP]::ToRoman($reading);
            }
        }
    }


}


function Invoke-SudachiTokenizer {
    <#
    .PARAMETER ignoreParen
    指定時は （） や ［］ に囲まれた部分に対する読み情報を付加しない
    .PARAMETER focusName
    指定時は2倍アキの後ろにあるノンブルや矢印後の見よ先項目は無視する
    .NOTES
    ビルドに rust を使用するようになったので、初回の pip install 時に rust がインストールされている必要がある。
    エラーメッセージで案内される https://rustup.rs/ をインストールして本体を再起動してから実行すれば解決する（はず）。
    #>

    param (
        [parameter(ValueFromPipeline)][string[]]$inputLine
        ,[switch]$ignoreParen
        ,[switch]$focusName
    )
    begin {
        $lines = New-Object System.Collections.ArrayList
    }
    process {
        $inputLine.ForEach({$lines.Add($_) > $null})
    }
    end {
        [SudachiPy]::new($lines, $ignoreParen, $focusName) | Write-Output
    }
}

function Get-ReadingWithSudachi {
    param (
        [parameter(ValueFromPipeline)][string[]]$inputLine
        ,[switch]$forBookIndex
    )
    begin {
        $lines = New-Object System.Collections.ArrayList
    }
    process {
        $inputLine.ForEach({
                $lines.Add($_) > $null
            })
    }
    end {
        $sudachi = [SudachiPy]::new($lines, $forBookIndex, $forBookIndex)
        $sudachi.GetReading() | Write-Output
    }
}

function Invoke-SortByReading {
    param (
        [parameter(ValueFromPipeline)][string[]]$inputLine
    )
    begin {
        $lines = New-Object System.Collections.ArrayList
    }
    process {
        $inputLine.ForEach({
                $lines.Add($_) > $null
            })
    }
    end {
        $sudachi = [SudachiPy]::new($lines, $forBookIndex, $forBookIndex)
        $sudachi.GetReading() | Sort-Object Normalized | ForEach-Object {$_.Line} | Write-Output
    }
}




function Convert-LinesToBookIndexReading {
    param (
        [switch]$asTsv
    )
    $result = $input | Get-ReadingWithSudachi -forBookIndex | Select-Object -Property "Reading", "Tokenize"
    if ($asTsv) {
        return $result | ForEach-Object {
            return $_.PSObject.Properties.Value | Join-String -Separator "`t"
        }
    }
    return $result
}

function uni {
    & ("C:\Users\{0}\Sync\portable_app\uni\uni.exe" -f $env:USERNAME) $args
}

function Get-LinesDelta {
    param(
        [int]$charsPerLine = 35,
        [int]$padding = 0,
        [string]$from
    )
    $after = $input | Join-String -Separator ""
    $afterGrids = [math]::Ceiling( [System.Text.Encoding]::GetEncoding("Shift_Jis").GetByteCount($after) / 2 )
    $beforeGrids = [math]::Ceiling( [System.Text.Encoding]::GetEncoding("Shift_Jis").GetByteCount($from) / 2 )
    $gridsDelta = [math]::Abs($beforeGrids - $afterGrids)
    $diff = "{0} -> {1}" -f $beforeGrids, $afterGrids
    if ($beforeGrids -lt $afterGrids) {
        if ($padding -lt $gridsDelta) {
            $linesDelta = [math]::Ceiling(($gridsDelta - $padding) / $charsPerLine) + 1
            return [PSCustomObject]@{
                "Grids" = "+{0} ({1})" -f $gridsDelta, $diff;
                "Lines" = "+{0}" -f $linesDelta;
            }
        }
        return [PSCustomObject]@{
            "Grids" = "+{0} ({1})" -f $gridsDelta, $diff;
            "Lines" = "+-0";
        }
    }
    $trailing = $charsPerLine - $padding
    if ($trailing -lt $gridsDelta) {
        $linesDelta = [math]::Floor(($gridsDelta - $trailing) / $charsPerLine) + 1
        return [PSCustomObject]@{
            "Grids" = "-{0} ({1})" -f $gridsDelta, $diff;
            "Lines" = "-{0}" -f $linesDelta;
        }
    }
    return [PSCustomObject]@{
        "Grids" = "-{0} ({1})" -f $gridsDelta, $diff;
        "Lines" = "+-0";
    }
}
Set-Alias linesDelta Get-LinesDelta

function Get-LinesSimilarityWithPython {
    $lines = New-Object System.Collections.ArrayList
    $input | Where-Object {$_.trim().length} | ForEach-Object {$lines.Add($_) > $null}

    $pyCodePath = $PSScriptRoot | Join-Path -ChildPath "python\get_similarity.py"
    Use-TempDir {
        $in = New-Item -Path ".\in.txt"
        $out = New-Item -Path ".\out.txt"
        $lines | Out-File -Encoding utf8NoBOM -FilePath $in.FullName
        Start-Process -Path python.exe -wait -ArgumentList @("-B", $pyCodePath, $in.FullName, $out.FullName) -NoNewWindow
        return Get-Content -Path $out.FullName | ConvertFrom-Json
    }
}

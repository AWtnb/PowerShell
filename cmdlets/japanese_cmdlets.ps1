
<# ==============================

cmdlets for processing japanese

            encoding: utf8bom
============================== #>

function aw {
    "あかさたなはまやらわ".GetEnumerator() | Write-Output
}

class DictRecord {
    [string]$reading
    [string]$word
    [string]$comment
    DictRecord([string]$reading, [string]$word, [string]$comment) {
        $this.reading = $reading
        $this.word = $word
        $this.comment = $comment
    }
    [string]GetLine([string]$pos) {
        return "{0}`t{1}`t{2}`t{3}" -f $this.reading, $this.word, $pos, $this.comment
    }
}

class HumanTsv {
    [string]$reading
    [string]$kana
    [string]$familyName
    [string]$restNameFull
    [string]$restNameCapitalized
    [string]$bio
    [string]$valiation
    [string]$keyword
    [string]$ref
    HumanTsv([string]$s) {
        $tsv = $s | ConvertFrom-Csv -Delimiter "`t" -Header "reading", "kana", "familyName", "restNameFull", "restNameCapitalized", "bio", "valiation", "keyword", "ref"
        $this.reading = $tsv.reading
        $this.kana = $tsv.kana
        $this.familyName = $tsv.familyName
        $this.restNameFull = $tsv.restNameFull
        $this.restNameCapitalized = $tsv.restNameCapitalized
        $this.bio = $tsv.bio
        $this.valiation = $tsv.valiation
        $this.keyword = $tsv.keyword
        $this.ref = $tsv.ref
    }
    [string]GetMain() {
        return [DictRecord]::new($this.reading, $this.kana, "").GetLine("人名")
    }
    [string]GetJpName() {
        # W. ジェームズ
        $word = $this.restNameCapitalized + " " + $this.kana
        return [DictRecord]::new($this.reading, $word, "日本語表記").GetLine("人名")
    }
    [string]GetJpNameReverse() {
        # ジェームズ， W.
        $word = $this.kana + "，" + $this.restNameCapitalized
        return [DictRecord]::new($this.reading, $word, "日本語出典表記").GetLine("人名")
    }
    [string]GetFullName() {
        # William James
        $word = $this.restNameFull + " " + $this.familyName
        return [DictRecord]::new($this.reading, $word, "欧文表記").GetLine("人名")
    }
    [string]GetFullNameReverse() {
        # James, William.
        $word = $this.familyName + ", " + $this.restNameFull
        return [DictRecord]::new($this.reading, $word, "欧文出典表記").GetLine("人名")
    }
    [string]GetName() {
        # W. James
        $word = $this.restNameCapitalized + " " + $this.familyName
        return [DictRecord]::new($this.reading, $word, "欧文短縮表記").GetLine("人名")
    }
    [string]GetNameReverse() {
        # James, W.
        $word = $this.familyName + ", " + $this.restNameCapitalized
        return [DictRecord]::new($this.reading, $word, "欧文出典短縮表記").GetLine("人名")
    }
    [string]GetDictContent() {
        # ジェームズ（James, William）
        $word = "{0}（{1}, {2}）" -f $this.kana, $this.familyName, $this.restNameFull
        if ($this.valiation.Length) {
            $comment = $this.valiation
        }
        else {
            $detail = @($this.keyword, $this.ref) | Where-Object {$_.length} | Join-String -Separator ";"
            $comment = "[{0}]{1}" -f $this.bio, $detail
            if ([System.Text.Encoding]::GetEncoding("Shift_Jis").GetByteCount($comment) -ge 250) {
                "more than 250bytes! : {0}" -f $word | Write-Host
            }
        }
        return [DictRecord]::new($this.reading, $word, $comment).GetLine("人名")
    }
}

function Convert-DictHumanTsvToImeSrc {
    $lines = New-Object System.Collections.ArrayList
    $input | ForEach-Object {
        $rec = [HumanTsv]::new($_)
        $lines.Add( $rec.GetMain() ) > $null
        $lines.Add( $rec.GetJpName() ) > $null
        $lines.Add( $rec.GetJpNameReverse() ) > $null
        $lines.Add( $rec.GetFullName() ) > $null
        $lines.Add( $rec.GetFullNameReverse() ) > $null
        $lines.Add( $rec.GetName() ) > $null
        $lines.Add( $rec.GetNameReverse() ) > $null
        $lines.Add( $rec.GetDictContent() ) > $null
    }
    $lines | Sort-Object -Unique | Out-File -FilePath "dict_human.txt" -Encoding utf8NoBOM -Force
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

    static [string] ToKanjiNum ([string]$strData) {
        $dict = @{
            "0" = "〇";
            "1" = "一";
            "2" = "二";
            "3" = "三";
            "4" = "四";
            "5" = "五";
            "6" = "六";
            "7" = "七";
            "8" = "八";
            "9" = "九";
        }
        $converted = [FormatJP]::ToHalfWidth($strData)
        $dict.GetEnumerator() | ForEach-Object {
            $converted = $converted.Replace($_.key, $_.value)
        }
        return $converted
    }

    static [string] FromRoman ([string]$strData) {
        $dict = @{
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
        }
        $converted = $strData
        $dict.GetEnumerator() | ForEach-Object {
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
    "ConvertTo-Katakana" = "ToKatakana";
    "ConvertTo-Hiragana" = "ToHiragana";
    "ConvertTo-Hairetsu" = "Normalize";
    "ConvertTo-KanjiNum" = "ToKanjiNum";
    "ConvertFrom-Roman" = "FromRoman";
    "ConvertTo-Roman" = "ToRoman";
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

Set-Alias toHan ConvertTo-HalfWidth
Set-Alias toHalf ConvertTo-HalfWidth

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
        [parameter(ValueFromPipeline = $true)]$code
    )
    begin {}
    process {
        [Unicode]::ToLetter($code) | Write-Output
    }
    end {}
}

function Convert-LetterToUnicode {
    param (
        [parameter(ValueFromPipeline = $true)]$s
    )
    begin {}
    process {
        $s.GetEnumerator() | ForEach-Object {
            [Unicode]::FromLetter($_) | Write-Output
        }
    }
    end {}
}

class SudachiTokenParser {
    [string]$reading = ""
    [string]$detail = ""
    [string]$markup = ""

    SudachiTokenParser($token) {
        $surface = $token.surface
        if ($token.pos -match "記号" -or $token.pos -match "空白" -or $surface -match "^([ァ-ヴ・ー]|[a-zA-Zａ-ｚＡ-Ｚ]|[0-9０-９]|[\W\s])+$") {
            $this.reading = $surface
            $this.detail = $surface
            $this.markup = $surface
            return
        }
        if (-not $token.reading) {
            $this.reading = $surface
            $this.detail = "{0}(?)" -f $surface
            $this.markup = $surface
            return
        }
        $this.reading = $token.reading
        $this.detail = ($surface -match "^[ぁ-ん]+$")? $surface : ("{0}({1})" -f $surface, $this.reading)
        $this.markup = ($surface -match "^[ぁ-ん]+$")? $surface : ("<ruby>{0}<rt>{1}</rt></ruby>" -f $surface, $this.reading)
    }

}

class SudachiTokensReader {
    [PSCustomObject[]]$tokens

    SudachiTokensReader($rawTokens) {
        $this.tokens = $rawTokens.ForEach({[SudachiTokenParser]::new($_)})
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

    [string] GetMarkup() {
        $stack = New-Object System.Collections.ArrayList
        foreach ($token in $this.tokens) {
            $stack.Add($token.markup) > $null
        }
        return $stack | Join-String -Separator "" -OutputPrefix "<p>" -OutputSuffix "</p>";
    }

}

class SudachiPy {
    [PSCustomObject[]]$parsed

    SudachiPy([string[]]$lines, [bool]$ignoreParen=$false) {
        $sudachiPath = $PSScriptRoot | Join-Path -ChildPath "python\sudachi_tokenizer.py"
        Use-TempDir {
            $in = New-Item -Path ".\in.txt"
            $out = New-Item -Path ".\out.txt"
            $lines | Out-File -Encoding utf8NoBOM -FilePath $in.FullName
            $opt = ($ignoreParen)? "IgnoreParen" : "IncludeParen"
            Start-Process -Path python.exe -wait -ArgumentList @("-B", $sudachiPath, $in.FullName, $out.FullName, $opt) -NoNewWindow
            $this.parsed = Get-Content -Path $out.FullName -Encoding utf8NoBOM | ConvertFrom-Json
        }
    }

    [PSCustomObject[]] GetReading() {
        return $this.parsed | ForEach-Object {
            $reader = [SudachiTokensReader]::new($_.tokens)
            $reading = $reader.GetReading()
            return [PSCustomObject]@{
                "Line" = $_.line;
                "Reading" = $reading;
                "Tokenize" = $reader.GetDetail();
                "Normalized" = [FormatJP]::Normalize($reading);
                "Roman" = [FormatJP]::ToRoman($reading);
                "Markup" = $reader.GetMarkup();
            }
        }
    }


}


function Invoke-SudachiTokenizer {
    <#
    .PARAMETER ignoreParen
    指定時は （） や ［］ に囲まれた部分に対する読み情報を付加しない
    .NOTES
    ビルド時にエラーが起きる場合は Build Tools for Visual Studio 2019 をインストールもしくはアップデートすること。
    https://github.com/WorksApplications/SudachiPy/issues/145
    https://sudachi-dev.slack.com/archives/CBCF278AC/p1604632785006000?thread_ts=1604632556.005900&cid=CBCF278AC

    #>

    param (
        [parameter(ValueFromPipeline = $true)][string[]]$inputLine
        ,[switch]$ignoreParen
    )
    begin {
        $lines = New-Object System.Collections.ArrayList
    }
    process {
        $inputLine.ForEach({$lines.Add($_) > $null})
    }
    end {
        [SudachiPy]::new($lines, $ignoreParen) | Write-Output
    }
}

function Get-ReadingWithSudachi {
    param (
        [parameter(ValueFromPipeline = $true)][string[]]$inputLine
        ,[switch]$forBookIndex
    )
    begin {
        $lines = New-Object System.Collections.ArrayList
    }
    process {
        $inputLine.ForEach({
            if ($forBookIndex) {
                $trim = $_ -replace "　　.+" -replace "　→.+"
                $lines.Add($trim) > $null
            }
            else {
                $lines.Add($_) > $null
            }
        })
    }
    end {
        $sudachi = [SudachiPy]::new($lines, $forBookIndex)
        $sudachi.GetReading() | Write-Output
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
    & ("C:\Users\{0}\Dropbox\portable_apps\cli\uni\uni.exe" -f $env:USERNAME) $args
}
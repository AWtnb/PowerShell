
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

Update-TypeData -MemberName "ToGodanReg" -TypeName System.String -Force -MemberType ScriptMethod -Value {
    $s = $this -as [string]
    $last = $s.Substring($s.Length - 1)
    if ($last -notmatch "[ぁ-ん]") {
        return $s
    }
    $pattern = switch -Regex ($last) {
        "[わいうえお]" {"[わいうえおっ]" ; break}
        "[かきくけこ]" {"[かきくけこい]" ; break}
        "[さしすせそ]" {"[さしすせそ]" ; break}
        "[なにぬねの]" {"[なにぬねのん]" ; break }
        "[まみむめも]" {"[まみむめもん]" ; break }
        "[らりるれろ]" {"[らりるれろ]" ; break }
        "[がぎぐげご]" {"[がぎぐげごい]" ; break }
        default {""}
    }
    if ($pattern.Length -lt 1) {
        return $s
    }
    return $s.Substring(0, 1) + $pattern
}


Update-TypeData -MemberName "ToNum" -TypeName System.String -Force -MemberType ScriptMethod -Value {
    $s = $this
    @{
        "一"= 1;
        "二"= 2;
        "三"= 3;
        "四"= 4;
        "五"= 5;
        "六"= 6;
        "七"= 7;
        "八"= 8;
        "九"= 9;
        "〇"= 0;
    }.GetEnumerator() | ForEach-Object {
        $s = $s -replace $_.Key, $_.Value
    }
    return $s
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

Update-TypeData -MemberName "ToUnicode" -TypeName System.String -Force -MemberType ScriptMethod -Value {
    return $this.GetEnumerator() | ForEach-Object { [Unicode]::FromLetter($_) }
}
Update-TypeData -MemberName "FromUnicode" -TypeName System.String -Force -MemberType ScriptMethod -Value {
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

class KanjiTable {
    [hashtable]$table

    # 常用漢字表のうち、ひらがなが一位に定まるもの（訓読みが1つしかないもの）
    $replacableJoyo = @{
        "哀" = "あわ";
        "悪" = "わる";
        "握" = "にぎ";
        "扱" = "あつか";
        "宛" = "あ";
        "嵐" = "あらし";
        "安" = "やす";
        "暗" = "くら";
        "衣" = "ころも";
        "位" = "くらい";
        "囲" = "かこ";
        "委" = "ゆだ";
        "畏" = "おそ";
        "異" = "こと";
        "移" = "うつ";
        "萎" = "な";
        "偉" = "えら";
        "違" = "ちが";
        "慰" = "なぐさ";
        "一" = "ひと";
        "茨" = "いばら";
        "芋" = "いも";
        "引" = "ひ";
        "印" = "しるし";
        "因" = "よ";
        "淫" = "みだ";
        "陰" = "かげ";
        "飲" = "の";
        "隠" = "かく";
        "右" = "みぎ";
        "唄" = "うた";
        "畝" = "うね";
        "浦" = "うら";
        "運" = "はこ";
        "雲" = "くも";
        "永" = "なが";
        "泳" = "およ";
        "営" = "いとな";
        "詠" = "よ";
        "影" = "かげ";
        "鋭" = "するど";
        "易" = "やさ";
        "越" = "こ";
        "円" = "まる";
        "延" = "の";
        "沿" = "そ";
        "炎" = "ほのお";
        "園" = "その";
        "猿" = "さる";
        "遠" = "とお";
        "鉛" = "なまり";
        "塩" = "しお";
        "縁" = "ふち";
        "艶" = "つや";
        "応" = "こた";
        "押" = "お";
        "殴" = "なぐ";
        "桜" = "さくら";
        "奥" = "おく";
        "横" = "よこ";
        "岡" = "おか";
        "屋" = "や";
        "虞" = "おそれ";
        "俺" = "おれ";
        "温" = "あたた";
        "穏" = "おだ";
        "化" = "ば";
        "加" = "くわ";
        "仮" = "かり";
        "花" = "はな";
        "価" = "あたい";
        "果" = "は";
        "河" = "かわ";
        "架" = "か";
        "夏" = "なつ";
        "荷" = "に";
        "華" = "はな";
        "渦" = "うず";
        "暇" = "ひま";
        "靴" = "くつ";
        "歌" = "うた";
        "稼" = "かせ";
        "蚊" = "か";
        "牙" = "きば";
        "瓦" = "かわら";
        "芽" = "め";
        "回" = "まわ";
        "灰" = "はい";
        "会" = "あ";
        "快" = "こころよ";
        "戒" = "いまし";
        "改" = "あらた";
        "怪" = "あや";
        "海" = "うみ";
        "皆" = "みな";
        "塊" = "かたまり";
        "解" = "と";
        "潰" = "つぶ";
        "壊" = "こわ";
        "貝" = "かい";
        "崖" = "がけ";
        "街" = "まち";
        "蓋" = "ふた";
        "垣" = "かき";
        "柿" = "かき";
        "各" = "おのおの";
        "革" = "かわ";
        "殻" = "から";
        "隔" = "へだ";
        "確" = "たし";
        "獲" = "え";
        "学" = "まな";
        "岳" = "たけ";
        "楽" = "たの";
        "額" = "ひたい";
        "顎" = "あご";
        "潟" = "かた";
        "渇" = "かわ";
        "葛" = "くず";
        "且" = "か";
        "株" = "かぶ";
        "釜" = "かま";
        "鎌" = "かま";
        "刈" = "か";
        "甘" = "あま";
        "汗" = "あせ";
        "肝" = "きも";
        "冠" = "かんむり";
        "乾" = "かわ";
        "患" = "わずら";
        "貫" = "つらぬ";
        "寒" = "さむ";
        "堪" = "た";
        "換" = "か";
        "勧" = "すす";
        "幹" = "みき";
        "慣" = "な";
        "管" = "くだ";
        "緩" = "ゆる";
        "館" = "やかた";
        "鑑" = "かんが";
        "丸" = "まる";
        "含" = "ふく";
        "岸" = "きし";
        "岩" = "いわ";
        "眼" = "まなこ";
        "顔" = "かお";
        "願" = "ねが";
        "企" = "くわだ";
        "机" = "つくえ";
        "忌" = "い";
        "祈" = "いの";
        "既" = "すで";
        "記" = "しる";
        "起" = "お";
        "飢" = "う";
        "鬼" = "おに";
        "帰" = "かえ";
        "寄" = "よ";
        "亀" = "かめ";
        "喜" = "よろこ";
        "幾" = "いく";
        "旗" = "はた";
        "器" = "うつわ";
        "輝" = "かがや";
        "機" = "はた";
        "技" = "わざ";
        "欺" = "あざむ";
        "疑" = "うたが";
        "戯" = "たわむ";
        "詰" = "つ";
        "脚" = "あし";
        "逆" = "さか";
        "虐" = "しいた";
        "九" = "ここの";
        "久" = "ひさ";
        "及" = "およ";
        "弓" = "ゆみ";
        "丘" = "おか";
        "休" = "やす";
        "吸" = "す";
        "朽" = "く";
        "臼" = "うす";
        "求" = "もと";
        "究" = "きわ";
        "泣" = "な";
        "急" = "いそ";
        "救" = "すく";
        "球" = "たま";
        "嗅" = "か";
        "窮" = "きわ";
        "牛" = "うし";
        "去" = "さ";
        "居" = "い";
        "拒" = "こば";
        "挙" = "あ";
        "許" = "ゆる";
        "御" = "おん";
        "共" = "とも";
        "叫" = "さけ";
        "狂" = "くる";
        "挟" = "はさ";
        "恐" = "おそ";
        "恭" = "うやうや";
        "境" = "さかい";
        "橋" = "はし";
        "矯" = "た";
        "鏡" = "かがみ";
        "響" = "ひび";
        "驚" = "おどろ";
        "暁" = "あかつき";
        "業" = "わざ";
        "凝" = "こ";
        "曲" = "ま";
        "極" = "きわ";
        "玉" = "たま";
        "近" = "ちか";
        "勤" = "つと";
        "琴" = "こと";
        "筋" = "すじ";
        "僅" = "わず";
        "錦" = "にしき";
        "謹" = "つつし";
        "襟" = "えり";
        "駆" = "か";
        "愚" = "おろ";
        "隅" = "すみ";
        "串" = "くし";
        "掘" = "ほ";
        "熊" = "くま";
        "繰" = "く";
        "君" = "きみ";
        "薫" = "かお";
        "兄" = "あに";
        "茎" = "くき";
        "型" = "かた";
        "契" = "ちぎ";
        "計" = "はか";
        "恵" = "めぐ";
        "掲" = "かか";
        "経" = "へ";
        "蛍" = "ほたる";
        "敬" = "うやま";
        "傾" = "かたむ";
        "携" = "たずさ";
        "継" = "つ";
        "詣" = "もう";
        "憩" = "いこ";
        "鶏" = "にわとり";
        "迎" = "むか";
        "鯨" = "くじら";
        "隙" = "すき";
        "撃" = "う";
        "激" = "はげ";
        "桁" = "けた";
        "欠" = "か";
        "穴" = "あな";
        "血" = "ち";
        "決" = "き";
        "潔" = "いさぎよ";
        "月" = "つき";
        "犬" = "いぬ";
        "見" = "み";
        "肩" = "かた";
        "建" = "た";
        "研" = "と";
        "兼" = "か";
        "剣" = "つるぎ";
        "拳" = "こぶし";
        "軒" = "のき";
        "健" = "すこ";
        "険" = "けわ";
        "堅" = "かた";
        "絹" = "きぬ";
        "遣" = "つか";
        "賢" = "かしこ";
        "鍵" = "かぎ";
        "繭" = "まゆ";
        "懸" = "か";
        "元" = "もと";
        "幻" = "まぼろし";
        "弦" = "つる";
        "限" = "かぎ";
        "原" = "はら";
        "現" = "あらわ";
        "減" = "へ";
        "源" = "みなもと";
        "己" = "おのれ";
        "戸" = "と";
        "古" = "ふる";
        "呼" = "よ";
        "固" = "かた";
        "股" = "また";
        "虎" = "とら";
        "故" = "ゆえ";
        "枯" = "か";
        "湖" = "みずうみ";
        "雇" = "やと";
        "誇" = "ほこ";
        "鼓" = "つづみ";
        "顧" = "かえり";
        "五" = "いつ";
        "互" = "たが";
        "悟" = "さと";
        "語" = "かた";
        "誤" = "あやま";
        "口" = "くち";
        "公" = "おおやけ";
        "巧" = "たく";
        "広" = "ひろ";
        "向" = "む";
        "江" = "え";
        "考" = "かんが";
        "攻" = "せ";
        "効" = "き";
        "厚" = "あつ";
        "候" = "そうろう";
        "耕" = "たがや";
        "貢" = "みつ";
        "高" = "たか";
        "控" = "ひか";
        "喉" = "のど";
        "慌" = "あわ";
        "港" = "みなと";
        "硬" = "かた";
        "溝" = "みぞ";
        "構" = "かま";
        "綱" = "つな";
        "興" = "おこ";
        "鋼" = "はがね";
        "乞" = "こ";
        "告" = "つ";
        "谷" = "たに";
        "刻" = "きざ";
        "国" = "くに";
        "黒" = "くろ";
        "骨" = "ほね";
        "駒" = "こま";
        "込" = "こ";
        "頃" = "ころ";
        "今" = "いま";
        "困" = "こま";
        "恨" = "うら";
        "根" = "ね";
        "痕" = "あと";
        "魂" = "たましい";
        "懇" = "ねんご";
        "左" = "ひだり";
        "砂" = "すな";
        "唆" = "そそのか";
        "差" = "さ";
        "鎖" = "くさり";
        "座" = "すわ";
        "再" = "ふたた";
        "災" = "わざわ";
        "妻" = "つま";
        "砕" = "くだ";
        "彩" = "いろど";
        "採" = "と";
        "済" = "す";
        "祭" = "まつ";
        "菜" = "な";
        "最" = "もっと";
        "催" = "もよお";
        "塞" = "ふさ";
        "載" = "の";
        "際" = "きわ";
        "埼" = "さい";
        "在" = "あ";
        "罪" = "つみ";
        "崎" = "さき";
        "作" = "つく";
        "削" = "けず";
        "酢" = "す";
        "搾" = "しぼ";
        "咲" = "さ";
        "札" = "ふだ";
        "刷" = "す";
        "撮" = "と";
        "擦" = "す";
        "皿" = "さら";
        "山" = "やま";
        "参" = "まい";
        "蚕" = "かいこ";
        "惨" = "みじ";
        "傘" = "かさ";
        "散" = "ち";
        "酸" = "す";
        "残" = "のこ";
        "斬" = "き";
        "子" = "こ";
        "支" = "ささ";
        "止" = "と";
        "氏" = "うじ";
        "仕" = "つか";
        "市" = "いち";
        "矢" = "や";
        "旨" = "むね";
        "死" = "し";
        "糸" = "いと";
        "至" = "いた";
        "伺" = "うかが";
        "使" = "つか";
        "刺" = "さ";
        "始" = "はじ";
        "姉" = "あね";
        "枝" = "えだ";
        "姿" = "すがた";
        "思" = "おも";
        "施" = "ほどこ";
        "紙" = "かみ";
        "脂" = "あぶら";
        "紫" = "むらさき";
        "歯" = "は";
        "飼" = "か";
        "賜" = "たまわ";
        "諮" = "はか";
        "示" = "しめ";
        "字" = "あざ";
        "寺" = "てら";
        "耳" = "みみ";
        "自" = "みずか";
        "似" = "に";
        "事" = "こと";
        "侍" = "さむらい";
        "持" = "も";
        "時" = "とき";
        "慈" = "いつく";
        "辞" = "や";
        "𠮟" = "しか";
        "失" = "うしな";
        "室" = "むろ";
        "執" = "と";
        "湿" = "しめ";
        "漆" = "うるし";
        "芝" = "しば";
        "写" = "うつ";
        "社" = "やしろ";
        "車" = "くるま";
        "者" = "もの";
        "射" = "い";
        "捨" = "す";
        "斜" = "なな";
        "煮" = "に";
        "遮" = "さえぎ";
        "謝" = "あやま";
        "蛇" = "へび";
        "借" = "か";
        "酌" = "く";
        "弱" = "よわ";
        "寂" = "さび";
        "取" = "と";
        "狩" = "か";
        "首" = "くび";
        "殊" = "こと";
        "腫" = "は";
        "種" = "たね";
        "趣" = "おもむき";
        "寿" = "ことぶき";
        "受" = "う";
        "呪" = "のろ";
        "授" = "さず";
        "収" = "おさ";
        "州" = "す";
        "秀" = "ひい";
        "周" = "まわ";
        "拾" = "ひろ";
        "秋" = "あき";
        "修" = "おさ";
        "袖" = "そで";
        "終" = "お";
        "習" = "なら";
        "就" = "つ";
        "愁" = "うれ";
        "醜" = "みにく";
        "蹴" = "け";
        "襲" = "おそ";
        "汁" = "しる";
        "充" = "あ";
        "住" = "す";
        "柔" = "やわ";
        "渋" = "しぶ";
        "獣" = "けもの";
        "縦" = "たて";
        "祝" = "いわ";
        "宿" = "やど";
        "縮" = "ちぢ";
        "熟" = "う";
        "述" = "の";
        "春" = "はる";
        "瞬" = "またた";
        "巡" = "めぐ";
        "盾" = "たて";
        "所" = "ところ";
        "書" = "か";
        "暑" = "あつ";
        "緒" = "お";
        "除" = "のぞ";
        "升" = "ます";
        "召" = "め";
        "招" = "まね";
        "承" = "うけたまわ";
        "昇" = "のぼ";
        "松" = "まつ";
        "沼" = "ぬま";
        "宵" = "よい";
        "唱" = "とな";
        "商" = "あきな";
        "焼" = "や";
        "詔" = "みことのり";
        "照" = "て";
        "詳" = "くわ";
        "障" = "さわ";
        "憧" = "あこが";
        "償" = "つぐな";
        "鐘" = "かね";
        "丈" = "たけ";
        "乗" = "の";
        "城" = "しろ";
        "情" = "なさ";
        "場" = "ば";
        "蒸" = "む";
        "縄" = "なわ";
        "譲" = "ゆず";
        "醸" = "かも";
        "色" = "いろ";
        "植" = "う";
        "殖" = "ふ";
        "飾" = "かざ";
        "織" = "お";
        "辱" = "はずかし";
        "尻" = "しり";
        "心" = "こころ";
        "申" = "もう";
        "伸" = "の";
        "身" = "み";
        "辛" = "から";
        "侵" = "おか";
        "津" = "つ";
        "唇" = "くちびる";
        "振" = "ふ";
        "浸" = "ひた";
        "真" = "ま";
        "針" = "はり";
        "深" = "ふか";
        "進" = "すす";
        "森" = "もり";
        "診" = "み";
        "寝" = "ね";
        "慎" = "つつし";
        "震" = "ふる";
        "薪" = "たきぎ";
        "人" = "ひと";
        "刃" = "は";
        "尽" = "つ";
        "甚" = "はなは";
        "尋" = "たず";
        "図" = "はか";
        "水" = "みず";
        "吹" = "ふ";
        "垂" = "た";
        "炊" = "た";
        "粋" = "いき";
        "衰" = "おとろ";
        "推" = "お";
        "酔" = "よ";
        "遂" = "と";
        "穂" = "ほ";
        "錘" = "つむ";
        "据" = "す";
        "杉" = "すぎ";
        "裾" = "すそ";
        "瀬" = "せ";
        "井" = "い";
        "世" = "よ";
        "成" = "な";
        "西" = "にし";
        "青" = "あお";
        "政" = "まつりごと";
        "星" = "ほし";
        "清" = "きよ";
        "婿" = "むこ";
        "晴" = "は";
        "勢" = "いきお";
        "誠" = "まこと";
        "誓" = "ちか";
        "静" = "しず";
        "整" = "ととの";
        "夕" = "ゆう";
        "赤" = "あか";
        "昔" = "むかし";
        "惜" = "お";
        "責" = "せ";
        "跡" = "あと";
        "積" = "つ";
        "切" = "き";
        "拙" = "つたな";
        "接" = "つ";
        "設" = "もう";
        "雪" = "ゆき";
        "節" = "ふし";
        "説" = "と";
        "舌" = "した";
        "絶" = "た";
        "千" = "ち";
        "川" = "かわ";
        "先" = "さき";
        "専" = "もっぱ";
        "泉" = "いずみ";
        "浅" = "あさ";
        "洗" = "あら";
        "扇" = "おうぎ";
        "煎" = "い";
        "羨" = "うらや";
        "銭" = "ぜに";
        "選" = "えら";
        "薦" = "すす";
        "鮮" = "あざ";
        "前" = "まえ";
        "善" = "よ";
        "繕" = "つくろ";
        "狙" = "ねら";
        "阻" = "はば";
        "粗" = "あら";
        "疎" = "うと";
        "訴" = "うった";
        "遡" = "さかのぼ";
        "礎" = "いしずえ";
        "双" = "ふた";
        "早" = "はや";
        "争" = "あらそ";
        "走" = "はし";
        "奏" = "かな";
        "相" = "あい";
        "草" = "くさ";
        "送" = "おく";
        "倉" = "くら";
        "捜" = "さが";
        "挿" = "さ";
        "桑" = "くわ";
        "巣" = "す";
        "掃" = "は";
        "爽" = "さわ";
        "窓" = "まど";
        "創" = "つく";
        "喪" = "も";
        "痩" = "や";
        "葬" = "ほうむ";
        "装" = "よそお";
        "遭" = "あ";
        "霜" = "しも";
        "騒" = "さわ";
        "藻" = "も";
        "造" = "つく";
        "憎" = "にく";
        "蔵" = "くら";
        "贈" = "おく";
        "束" = "たば";
        "促" = "うなが";
        "息" = "いき";
        "捉" = "とら";
        "側" = "がわ";
        "測" = "はか";
        "続" = "つづ";
        "率" = "ひき";
        "村" = "むら";
        "孫" = "まご";
        "損" = "そこ";
        "他" = "ほか";
        "多" = "おお";
        "打" = "う";
        "唾" = "つば";
        "太" = "ふと";
        "体" = "からだ";
        "耐" = "た";
        "待" = "ま";
        "退" = "しりぞ";
        "袋" = "ふくろ";
        "替" = "か";
        "貸" = "か";
        "滞" = "とどこお";
        "大" = "おお";
        "滝" = "たき";
        "沢" = "さわ";
        "濁" = "にご";
        "但" = "ただ";
        "脱" = "ぬ";
        "奪" = "うば";
        "棚" = "たな";
        "誰" = "だれ";
        "炭" = "すみ";
        "淡" = "あわ";
        "短" = "みじか";
        "嘆" = "なげ";
        "綻" = "ほころ";
        "鍛" = "きた";
        "男" = "おとこ";
        "暖" = "あたた";
        "池" = "いけ";
        "知" = "し";
        "致" = "いた";
        "置" = "お";
        "竹" = "たけ";
        "蓄" = "たくわ";
        "築" = "きず";
        "中" = "なか";
        "仲" = "なか";
        "虫" = "むし";
        "沖" = "おき";
        "注" = "そそ";
        "昼" = "ひる";
        "柱" = "はしら";
        "鋳" = "い";
        "弔" = "とむら";
        "兆" = "きざ";
        "町" = "まち";
        "長" = "なが";
        "挑" = "いど";
        "張" = "は";
        "彫" = "ほ";
        "眺" = "なが";
        "釣" = "つ";
        "鳥" = "とり";
        "朝" = "あさ";
        "貼" = "は";
        "超" = "こ";
        "嘲" = "あざけ";
        "潮" = "しお";
        "澄" = "す";
        "聴" = "き";
        "懲" = "こ";
        "沈" = "しず";
        "珍" = "めずら";
        "鎮" = "しず";
        "追" = "お";
        "痛" = "いた";
        "塚" = "つか";
        "漬" = "つ";
        "坪" = "つぼ";
        "鶴" = "つる";
        "低" = "ひく";
        "定" = "さだ";
        "底" = "そこ";
        "庭" = "にわ";
        "堤" = "つつみ";
        "提" = "さ";
        "程" = "ほど";
        "締" = "し";
        "諦" = "あきら";
        "泥" = "どろ";
        "的" = "まと";
        "笛" = "ふえ";
        "摘" = "つ";
        "敵" = "かたき";
        "溺" = "おぼ";
        "店" = "みせ";
        "添" = "そ";
        "転" = "ころ";
        "田" = "た";
        "伝" = "つた";
        "吐" = "は";
        "妬" = "ねた";
        "都" = "みやこ";
        "渡" = "わた";
        "塗" = "ぬ";
        "賭" = "か";
        "土" = "つち";
        "努" = "つと";
        "刀" = "かたな";
        "冬" = "ふゆ";
        "灯" = "ひ";
        "当" = "あ";
        "投" = "な";
        "豆" = "まめ";
        "東" = "ひがし";
        "倒" = "たお";
        "唐" = "から";
        "島" = "しま";
        "桃" = "もも";
        "討" = "う";
        "透" = "す";
        "悼" = "いた";
        "盗" = "ぬす";
        "湯" = "ゆ";
        "登" = "のぼ";
        "答" = "こた";
        "等" = "ひと";
        "筒" = "つつ";
        "統" = "す";
        "踏" = "ふ";
        "藤" = "ふじ";
        "闘" = "たたか";
        "同" = "おな";
        "洞" = "ほら";
        "動" = "うご";
        "童" = "わらべ";
        "道" = "みち";
        "働" = "はたら";
        "導" = "みちび";
        "瞳" = "ひとみ";
        "峠" = "とうげ";
        "独" = "ひと";
        "栃" = "とち";
        "突" = "つ";
        "届" = "とど";
        "豚" = "ぶた";
        "貪" = "むさぼ";
        "鈍" = "にぶ";
        "曇" = "くも";
        "内" = "うち";
        "梨" = "なし";
        "謎" = "なぞ";
        "鍋" = "なべ";
        "南" = "みなみ";
        "軟" = "やわ";
        "二" = "ふた";
        "尼" = "あま";
        "匂" = "にお";
        "虹" = "にじ";
        "任" = "まか";
        "忍" = "しの";
        "認" = "みと";
        "熱" = "あつ";
        "年" = "とし";
        "粘" = "ねば";
        "燃" = "も";
        "悩" = "なや";
        "濃" = "こ";
        "波" = "なみ";
        "破" = "やぶ";
        "罵" = "ののし";
        "拝" = "おが";
        "杯" = "さかずき";
        "配" = "くば";
        "敗" = "やぶ";
        "廃" = "すた";
        "売" = "う";
        "梅" = "うめ";
        "培" = "つちか";
        "買" = "か";
        "泊" = "と";
        "迫" = "せま";
        "剝" = "は";
        "薄" = "うす";
        "麦" = "むぎ";
        "縛" = "しば";
        "箱" = "はこ";
        "箸" = "はし";
        "肌" = "はだ";
        "髪" = "かみ";
        "抜" = "ぬ";
        "半" = "なか";
        "犯" = "おか";
        "帆" = "ほ";
        "伴" = "ともな";
        "坂" = "さか";
        "板" = "いた";
        "飯" = "めし";
        "煩" = "わずら";
        "比" = "くら";
        "皮" = "かわ";
        "否" = "いな";
        "卑" = "いや";
        "飛" = "と";
        "疲" = "つか";
        "秘" = "ひ";
        "被" = "こうむ";
        "悲" = "かな";
        "扉" = "とびら";
        "費" = "つい";
        "避" = "さ";
        "尾" = "お";
        "眉" = "まゆ";
        "美" = "うつく";
        "備" = "そな";
        "鼻" = "はな";
        "膝" = "ひざ";
        "肘" = "ひじ";
        "匹" = "ひき";
        "必" = "かなら";
        "筆" = "ふで";
        "姫" = "ひめ";
        "俵" = "たわら";
        "漂" = "ただよ";
        "猫" = "ねこ";
        "品" = "しな";
        "浜" = "はま";
        "貧" = "まず";
        "夫" = "おっと";
        "父" = "ちち";
        "付" = "つ";
        "布" = "ぬの";
        "怖" = "こわ";
        "赴" = "おもむ";
        "浮" = "う";
        "腐" = "くさ";
        "敷" = "し";
        "侮" = "あなど";
        "伏" = "ふ";
        "幅" = "はば";
        "腹" = "はら";
        "払" = "はら";
        "沸" = "わ";
        "仏" = "ほとけ";
        "物" = "もの";
        "紛" = "まぎ";
        "噴" = "ふ";
        "憤" = "いきどお";
        "奮" = "ふる";
        "文" = "ふみ";
        "聞" = "き";
        "併" = "あわ";
        "餅" = "もち";
        "米" = "こめ";
        "壁" = "かべ";
        "癖" = "くせ";
        "別" = "わか";
        "蔑" = "さげす";
        "片" = "かた";
        "返" = "かえ";
        "変" = "か";
        "偏" = "かたよ";
        "編" = "あ";
        "便" = "たよ";
        "保" = "たも";
        "補" = "おぎな";
        "母" = "はは";
        "募" = "つの";
        "墓" = "はか";
        "慕" = "した";
        "暮" = "く";
        "方" = "かた";
        "包" = "つつ";
        "芳" = "かんば";
        "奉" = "たてまつ";
        "宝" = "たから";
        "泡" = "あわ";
        "倣" = "なら";
        "峰" = "みね";
        "崩" = "くず";
        "報" = "むく";
        "蜂" = "はち";
        "豊" = "ゆた";
        "飽" = "あ";
        "褒" = "ほ";
        "縫" = "ぬ";
        "亡" = "な";
        "乏" = "とぼ";
        "忙" = "いそが";
        "妨" = "さまた";
        "忘" = "わす";
        "防" = "ふせ";
        "房" = "ふさ";
        "冒" = "おか";
        "紡" = "つむ";
        "望" = "のぞ";
        "傍" = "かたわ";
        "暴" = "あば";
        "膨" = "ふく";
        "謀" = "はか";
        "頰" = "ほお";
        "北" = "きた";
        "牧" = "まき";
        "墨" = "すみ";
        "堀" = "ほり";
        "本" = "もと";
        "翻" = "ひるがえ";
        "麻" = "あさ";
        "磨" = "みが";
        "妹" = "いもうと";
        "埋" = "う";
        "枕" = "まくら";
        "又" = "また";
        "末" = "すえ";
        "満" = "み";
        "味" = "あじ";
        "岬" = "みさき";
        "民" = "たみ";
        "眠" = "ねむ";
        "矛" = "ほこ";
        "務" = "つと";
        "無" = "な";
        "夢" = "ゆめ";
        "霧" = "きり";
        "娘" = "むすめ";
        "名" = "な";
        "命" = "いのち";
        "迷" = "まよ";
        "鳴" = "な";
        "滅" = "ほろ";
        "免" = "まぬか";
        "綿" = "わた";
        "茂" = "しげ";
        "毛" = "け";
        "網" = "あみ";
        "黙" = "だま";
        "門" = "かど";
        "匁" = "もんめ";
        "野" = "の";
        "弥" = "や";
        "訳" = "わけ";
        "薬" = "くすり";
        "躍" = "おど";
        "闇" = "やみ";
        "油" = "あぶら";
        "諭" = "さと";
        "癒" = "い";
        "友" = "とも";
        "有" = "あ";
        "勇" = "いさ";
        "湧" = "わ";
        "遊" = "あそ";
        "誘" = "さそ";
        "与" = "あた";
        "余" = "あま";
        "誉" = "ほま";
        "預" = "あず";
        "幼" = "おさな";
        "用" = "もち";
        "羊" = "ひつじ";
        "妖" = "あや";
        "揚" = "あ";
        "揺" = "ゆ";
        "葉" = "は";
        "溶" = "と";
        "腰" = "こし";
        "様" = "さま";
        "踊" = "おど";
        "窯" = "かま";
        "養" = "やしな";
        "抑" = "おさ";
        "浴" = "あ";
        "翼" = "つばさ";
        "裸" = "はだか";
        "雷" = "かみなり";
        "絡" = "から";
        "落" = "お";
        "乱" = "みだ";
        "卵" = "たまご";
        "藍" = "あい";
        "利" = "き";
        "里" = "さと";
        "裏" = "うら";
        "履" = "は";
        "離" = "はな";
        "立" = "た";
        "柳" = "やなぎ";
        "流" = "なが";
        "留" = "と";
        "竜" = "たつ";
        "粒" = "つぶ";
        "旅" = "たび";
        "良" = "よ";
        "涼" = "すず";
        "陵" = "みささぎ";
        "量" = "はか";
        "糧" = "かて";
        "力" = "ちから";
        "緑" = "みどり";
        "林" = "はやし";
        "輪" = "わ";
        "臨" = "のぞ";
        "涙" = "なみだ";
        "類" = "たぐ";
        "励" = "はげ";
        "戻" = "もど";
        "例" = "たと";
        "鈴" = "すず";
        "霊" = "たま";
        "麗" = "うるわ";
        "暦" = "こよみ";
        "劣" = "おと";
        "裂" = "さ";
        "練" = "ね";
        "路" = "じ";
        "露" = "つゆ";
        "弄" = "もてあそ";
        "朗" = "ほが";
        "漏" = "も";
        "麓" = "ふもと";
        "賄" = "まかな";
        "脇" = "わき";
        "惑" = "まど";
        "枠" = "わく";
        "腕" = "うで";
    }

    KanjiTable() {
        $this.table = $this.replacableJoyo
    }

    [void] Update([hashtable]$additional) {
        $additional.GetEnumerator() | ForEach-Object {
            $this.table[$_.Key] = $_.Value
        }
    }

    [string] Replace([string]$s) {
        $f = $s
        $this.table.GetEnumerator() | ForEach-Object {
            $f = $f -replace $_.Key, $_.Value
        }
        return $f
    }
}

Update-TypeData -MemberName "Hira" -TypeName System.String -Force -MemberType ScriptMethod -Value {
    $kTable = [KanjiTable]::new()
    $kTable.Update(@{
        "言" = "い";
        "上" = "うえ";
        "掛" = "か";
        "関" = "かか";
        "極" = "きわ";
        "事" = "こと";
        "様々" = "さまざま";
        "更" = "さら";
        "達" = "たち";
        "例" = "たと";
        "時" = "とき";
        "特" = "とく";
        "共" = "とも";
        "中" = "なか";
        "等" = "など";
        "何" = "なん";
        "方" = "ほう";
    })
    return $kTable.Replace($this)
}


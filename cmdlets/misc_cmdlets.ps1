
<# ==============================

MISC

                encoding: utf8bom
============================== #>


class Base64 {

    static [string] Encode ([string]$s) {
        $byte = ([System.Text.Encoding]::Default).GetBytes($s)
        return [Convert]::ToBase64String($byte)
    }

    static [string] Decode ([string]$s) {
        $byte = [System.Convert]::FromBase64String($s)
        return [System.Text.Encoding]::Default.GetString($byte)
    }

}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName ToBase64 -Value {
    return [Base64]::Encode($this)
}
Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName FromBase64 -Value {
    return [Base64]::Decode($this)
}

function ConvertFrom-Base64 {
    param (
        [parameter(ValueFromPipeline)][string]$inputLine
    )
    begin {}
    process {
        [Base64]::Decode($inputLine) | Write-Output
    }
    end {}
}
function ConvertTo-Base64 {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj -ErrorAction SilentlyContinue
        if ($fileObj) {
            $bytes = $fileObj | Get-Content -AsByteStream
            [PSCustomObject]@{
                "Name" = $fileObj.Name;
                "Encode" = [System.Convert]::ToBase64String($bytes);
            } | Write-Output
        }
        else {
            [PSCustomObject]@{
                "Name" = $inputObj;
                "Encode" = ($inputObj -as [string]).ToBase64();
                "Markup" = "<img src=`"data:image/png;base64,$()`""
            } | Write-Output
        }
    }
    end {}
}

class Base26 {

    static [int] ToDecimal ([string]$strData) {
        if ($strData -notmatch "^[a-z]+$") {
            return 0
        }
        $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        $strData = $strData.ToUpper()
        $ret = 0
        $strData.GetEnumerator() | ForEach-Object {
            $ret = $ret * 26 + $alphabet.IndexOf($_) + 1
        }
        return $ret
    }

    static [string] FromDecimal ([int]$intData) {
        if ($intData -le 0) {
            return ""
        }
        if ($intData -le 26) {
            return $([char]($intData + 64)).ToString()
        }
        $alphabetIndex = ($intData % 26)? $intData % 26 : 26
        $nCycle = [math]::Floor(($intData - $alphabetIndex) / 26)
        return $("{0}{1}" -f [Base26]::FromDecimal($nCycle), ([char]($alphabetIndex + 64)).ToString())
    }

}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName ToBase26 -Value {
    return [Base26]::ToDecimal($this)
}
Update-TypeData -TypeName "System.Int32" -Force -MemberType ScriptMethod -MemberName ToBase26 -Value {
    return [Base26]::FromDecimal($this)
}

function ConvertFrom-Base26 {
    param (
        [parameter(ValueFromPipeline)]$s
    )
    begin {}
    process {
        [Base26]::ToDecimal($s) | Write-Output
    }
    end {}
}

function ConvertTo-Base26 {
    param (
        [parameter(ValueFromPipeline)]$s
    )
    begin {}
    process {
        [Base26]::FromDecimal($s) | Write-Output
    }
    end {}
}

class RGB {

    static [int] ToInt ([int]$r, [int]$g, [int]$b) {
        return $r + $g*256 + $b*256*256
    }

    static [int[]] FromInt ([int]$i) {
        return [RGB]::FromColorcode([Colorcode]::FromInt($i))
    }

    static [string] ToColorcode ([int]$r, [int]$g, [int]$b) {
        return [Colorcode]::FromInt([RGB]::ToInt($r, $g, $b))
    }

    static [int[]] FromColorcode ([string]$colorcode) {
        if ($colorcode.StartsWith("#")) {
            $colorcode = $colorcode.TrimStart("#")
        }
        $r = [convert]::ToInt32($colorcode.Substring(0, 2), 16)
        $g = [convert]::ToInt32($colorcode.Substring(2, 2), 16)
        $b = [convert]::ToInt32($colorcode.Substring(4, 2), 16)
        return @($r, $g, $b)
    }


}

class Colorcode {

    static [int] ToInt ([string]$colorcode) {
        $rgb = [RGB]::FromColorcode($colorcode)
        return [RGB]::ToInt($rgb[0], $rgb[1], $rgb[2])
    }

    static [string] FromInt ([int]$i) {
        $r = $i % 256
        $g = (($i - $r) / 256) % 256
        $b = ($i - $r - $g*256) / (256*256)
        return "#{0:X2}{1:X2}{2:X2}" -f $r, $g, $b
    }

    static [int[]] ToRgb ([string]$colorcode) {
        return [RGB]::FromInt([Colorcode]::ToInt($colorcode))
    }

    static [string] FromRgb ([int]$r, [int]$g, [int]$b) {
        return [Colorcode]::FromInt([RGB]::ToInt($r, $g, $b))
    }



}

function Convert-Rgb2Int {
    param (
        [int]$r, [int]$g, [int]$b
    )
    [RGB]::ToInt($r, $g, $b) | Write-Output
}
function Convert-Rgb2Colorcode {
    param (
        [int]$r, [int]$g, [int]$b
    )
    [RGB]::ToColorcode($r, $g, $b) | Write-Output
}
function Convert-Int2Colorcode {
    param (
        [parameter(ValueFromPipeline)][int]$i
    )
    begin {}
    process {
        [Colorcode]::FromInt($i) | Write-Output
    }
    end {}
}
function Convert-Colorcode2Rgb {
    param (
        [parameter(ValueFromPipeline)]$colorcode
    )
    begin {}
    process {
        [Colorcode]::ToRgb($colorcode) | Write-Output
    }
    end {}
}
function Convert-Colorcode2Int {
    param (
        [parameter(ValueFromPipeline)]$colorcode
    )
    begin {}
    process {
        [Colorcode]::ToInt($colorcode) | Write-Output
    }
    end {}
}

function _alias {
    [OutputType("System.Management.Automation.AliasInfo")]
    param ([string]$name = ".", [string]$definition = ".")
    return $(Get-Alias | Where-Object Name -Match $name | Where-Object Definition -Match $definition)
}

function _cmdlet {
    [OutputType("System.Management.Automation.AliasInfo", "System.Management.Automation.FunctionInfo", "System.Management.Automation.CmdletInfo")]
    param ([string]$name)
    return $(Get-Command | Where-Object Name -Match $name)
}

function ConvertFrom-CsvUtil {
    param (
        [string]$delimiter = ","
        ,[string[]]$header
        ,[switch]$verbatim
    )
    $content = New-Object System.Collections.ArrayList
    $input.ForEach({
        $content.Add($_) > $null
    })
    if ($header -and $verbatim) {
        Write-Error "Only one of '-header' or '-verbatim' should be specified."
        return
    }
    if ($header) {
        return $($content | ConvertFrom-Csv -Delimiter $delimiter -Header $header)
    }
    if ($verbatim) {
        return $($content | ConvertFrom-Csv -Delimiter $delimiter)
    }
    $firstLine = ($content | Select-Object -First 1) -split $delimiter
    $header = 1..$firstLine.Count | ForEach-Object {
        return [Base26]::FromDecimal($_)
    }
    return $($content | ConvertFrom-Csv -Delimiter $delimiter -Header $header)
}
Set-Alias fromCSV ConvertFrom-CsvUtil

function fromTSV {
    param (
        [string[]]$header
        ,[switch]$verbatim
    )
    return $($input | ConvertFrom-CsvUtil -delimiter "`t" -header $header -verbatim:$verbatim)
}

function ConvertTo-Tsv {
    <#
        .EXAMPLE
        $hoge | toTSV
    #>
    param (
        [switch]$withHeader
    )
    $tsv = $input | ConvertTo-Csv -Delimiter "`t" -NoTypeInformation -UseQuotes Always
    return ($withHeader)? $tsv : $($tsv | Select-Object -Skip 1)
}
Set-Alias toTSV ConvertTo-Tsv
Set-Alias toCSV ConvertTo-Csv

function Edit-Tsv {
    param (
        [int]$rowIndex,
        [scriptblock]$formatBlock = {$_}
    )
    foreach ($line in $input) {
        $rows = $line -split "`t"
        $arr = New-Object System.Collections.ArrayList
        for ($r = 0; $r -lt $rows.Count; $r++) {
            if ($r -eq $rowIndex) {
                $arr.Add($($rows[$r] | ForEach-Object $formatBlock)) > $null
                continue
            }
            $arr.Add($rows[$r]) > $null
        }
        $arr -join "`t" | Write-Output
    }
}

function Get-DuplicateItem {
    $input | Group-Object | Where-Object Count -gt 1 | Sort-Object Count -Stable | Select-Object Name, Count | Write-Output
}
Set-Alias dupl Get-DuplicateItem

function Get-UniqueOrderdArray {
    $stack = New-Object System.Collections.ArrayList
    foreach ($i in $input) {
        if ($i -cin $stack) {
            continue
        }
        Write-Output $i
        $stack.Add($i) > $null
    }
}
Set-Alias uq Get-UniqueOrderdArray


if(-not ('Pwsh.SHA256' -as [type]))
{
}

function Find-SameFile {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$searchIn = "."
    )
    begin {
        $hashGruop = @{}
        Get-ChildItem -Path $searchIn -File -Recurse | Get-FileHash | Group-Object -Property Hash | ForEach-Object {
            $hashGruop.Add($_.Name, @($_.Group.Path))
        }
    }
    process {
        $target = Get-Item -LiteralPath $inputObj
        if ($target.Extension) {
            $targetHash = (Get-FileHash $target.Fullname).hash
            $found = New-Object System.Collections.ArrayList
            foreach ($p in $hashGruop[$targetHash]) {
                if ($p -and $p -ne $target.FullName) {
                    $found.Add( [System.IO.Path]::GetRelativePath($searchIn, $p) ) > $null
                }
            }
            return [PSCustomObject]@{
                "Path" = $($target.FullName | Resolve-Path -Relative);
                "Count" = $found.Count;
                "Found" = $found;
            }
        }
    }
    end {}
}

function Group-SameFile {
    <#
    .EXAMPLE
      ls | Get-FileHash | Group-SameFile
    #>
    $group = $input | Group-Object -Property Hash
    $group | ForEach-Object {
        if ($_.Count -gt 1) {
            $paths = $_.Group.Path | Sort-Object {$_.Length}
            $shortest = $paths | Select-Object -First 1
            $clones = @($paths | Select-Object -Skip 1 | Get-Item)
            return [PSCustomObject]@{
                "Shortest" = $shortest;
                "Clones" = $clones;
            }
        }
        return [PSCustomObject]@{
            "Shortest" = $_.Group.Path;
            "Clones" = @();
        }
    } | Sort-Object {$_.Shortest} | Write-Output
}

function Invoke-DiffStringArray {
    <#
        .EXAMPLE
        (ls).name | diffStrArray -deltaTo $hoge
    #>
    param (
        [string[]]$deltaTo,
        [switch]$asObject
    )

    $lines = New-Object System.Collections.ArrayList
    $input.ForEach({$lines.Add($_) > $null})

    $diff =  Compare-Object $lines $deltaTo
    if(-not $diff.Count) {
        Write-Host "same array!" -ForegroundColor Cyan
        return
    }

    $hashTable = @{}
    $diff | ForEach-Object {
        if ($_.SideIndicator -eq "=>") {
            $delta = 1
        }
        elseif ($_.SideIndicator -eq "<=") {
            $delta = -1
        }

        if ($_.InputObject -in $hashTable.Keys) {
            $hashTable[$_.InputObject] += $delta
        }
        else {
            $hashTable.Add($_.InputObject, $delta)
        }
    }

    $objArray = New-Object System.Collections.ArrayList
    $hashTable.GetEnumerator() | ForEach-Object {
        $record = [PSCustomObject]@{
            delta = $_.Value;
            line = $_.Key;
        }
        $objArray.Add($record) > $null
    }
    if ($asObject) {
        return $objArray
    }

    $objArray | Sort-Object line, delta | ForEach-Object {
        $color = ($_.delta -ge 1)? "Green": "Magenta"
        Write-Host ("[{0,2}]" -f $_.delta) -BackgroundColor $color -ForegroundColor Black -NoNewline
        Write-Host $_.line -ForegroundColor $color
    }
}
Set-Alias diffStrArray Invoke-DiffStringArray


function Invoke-DiffAsHtml {
    param (
        [parameter(Mandatory)][string]$from
        ,[parameter(Mandatory)][string]$to
        ,[string]$outName = "out"
        ,[switch]$compress
    )
    $outName = ($outName.EndsWith(".html"))? $outName : $outName + ".html"
    $fromPath = Resolve-Path -Path $from
    $toPath = Resolve-Path -Path $to
    $outPath = $pwd.Path | Join-Path -ChildPath $outName
    $pyCodePath = $PSScriptRoot | Join-Path -ChildPath "python\diff_as_html\diff.py"
    $cmd = 'python -B "{0}" "{1}" "{2}" "{3}"' -f $pyCodePath, $fromPath, $toPath, $outPath
    if ($compress) {
        $cmd += " --compress"
    }
    $cmd | Invoke-Expression
}

function Invoke-RecycleBin {
    <#
        .EXAMPLE
        ls | ? Name -match "hoge" | Invoke-RecycleBin
    #>
    $target = New-Object System.Collections.ArrayList
    $input | Where-Object {$_.GetType().Name -in @("DirectoryInfo", "FileInfo")} | ForEach-Object {$target.Add($_) > $null}
    if (-not $target) {
        Start-Process shell:RecycleBinFolder
        return
    }
    $counter = 0
    $target | ForEach-Object {
        $fullPath = $_.Fullname
        $fName = $_.Name
        try {
            $fullPath | Resolve-Path -Relative | Write-Host -ForegroundColor DarkBlue
            if ($_.GetType().Name -eq "DirectoryInfo") {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($fullPath, "OnlyErrorDialogs", "SendToRecycleBin")
            }
            elseif ($_.GetType().Name -eq "FileInfo") {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($fullPath, "OnlyErrorDialogs", "SendToRecycleBin")
            }
            $counter += 1
        }
        catch {
            "ERROR: failed to move '{0}' to recyclebin!" -f $fName | Write-Error
        }
    }
    "Recycled {0} items." -f $counter | Write-Host -ForegroundColor Cyan
}
Set-Alias gomi Invoke-RecycleBin

function ConvertTo-HashTableFromStringArray {
    <#
        .SYNOPSIS
        配列を指定文字で分割して連想配列化する
        .DESCRIPTION
        パイプライン経由での入力にのみ対応
        各行に区切り文字が2つ以上ある場合は最初の区切り文字で分割
        .PARAMETER delimiter
        区切り文字
        .EXAMPLE
        cat \.hoge.txt | ConvertTo-HashTableFromStringArray
    #>
    param (
        [string]$delimiter = "`t"
    )
    $hashTable = [ordered]@{}
    $input | ForEach-Object {
        $pair = @($_ -split $delimiter)
        $key = $pair[0]
        $value = $pair[1..($pair.Count - 1)] -join " "
        if ($key -in $hashTable.keys) {
            Write-Host ("ERROR: key '{0}' already has value '{1}'!" -f $key, $hashTable[$key]) -ForegroundColor Magenta
        }
        else {
            $hashTable[$key] = $value
        }
    }
    return $hashTable
}

# function ConvertTo-HashTableFromObject {
#     param (
#         [string]$keyProperty,
#         [string]$valueProperty
#     )
#     $target = $input | Select-Object -Property $keyProperty, $valueProperty
#     $duplicate = $target |Select-Object -ExpandProperty $keyProperty | Group-Object | Where-Object Count -gt 1
#     if ($duplicate) {
#         Write-Host "ERROR: duplicate key." -ForegroundColor Magenta
#         $duplicate | ForEach-Object {
#             Write-Host ("  {0} => {1}times." -f $_.Name,$_.Count) -ForegroundColor Magenta
#         }
#         return
#     }
#     $hashTable = [ordered]@{}
#     $target | ForEach-Object {
#         $hashTable[($_ | Select-Object -ExpandProperty $keyProperty)] = $_ | Select-Object -ExpandProperty $valueProperty
#     }
#     return $hashTable
# }

function Get-AddressByPostalCode {
    <#
        .SYNOPSIS
        thanks: https://github.com/madefor/postal-code-api/
        .EXAMPLE
        Get-AddressByPostalCode 1010051 # => 東京都千代田区神田神保町
        .EXAMPLE
        echo 1010051 1010052 1010053 | Get-AddressByPostalCode # => 東京都千代田区神田神保町
                                                            # => 東京都千代田区神田小川町
                                                            # => 東京都千代田区神田美土代町
    #>
    param (
        [parameter(ValueFromPipeline)]$inputLine
    )
    begin {
    }
    process {
        $postalCode = $inputLine -replace "[^\d]"
        if (-not $postalCode) {
            return ""
        }
        $filled = "{0:d7}" -f [int]$postalCode
        $url = "https://madefor.github.io/postal-code-api/api/v1/{0}/{1}.json" -f $filled.Substring(0,3), $filled.Substring(3,4)
        try {
            $res = Invoke-RestMethod -Uri $url -ErrorAction Stop
            return ($res.data.ja.psobject.properties.value -join "")
        }
        catch {
            return ""
        }
    }
    end {
    }
}



function colorChecker {
    $colors = @("White", "Black", "Blue", "DarkBlue", "Green", "DarkGreen", "Cyan", "DarkCyan", "Red", "DarkRed", "Magenta", "DarkMagenta", "Yellow", "DarkYellow", "Gray", "DarkGray")
    $colors | ForEach-Object {
        $back = $_
        Write-Host "background: [$($back)]"
        $colors | ForEach-Object {
            if ($_ -ne $back) {
                ("    $($_)").padRight(15) | Write-Host -BackgroundColor $back -ForegroundColor $_
            }
        }
        Write-Host ""
    }
}

function ANSICOL ([switch]$bg){
    # to enable 256 color in Cmder, need to run "AnsiColors256.ans"
    # cmd /c type "(path to AnsiColors256.ans)"
    for ($i = 0; $i -lt 16; $i++) {
        for ($j = 0; $j -lt 16; $j++) {
            $v = [int]$i * 16 + [int]$j
            $ansi = ($bg)? 4 : 3
            "`e[{0}8;5;{1}m{1:000}`e[0m " -f $ansi, $v | Write-Host -NoNewline
        }
        Write-Host
    }
}

function len ([switch]$net){
    $s = @($input) -join ""
    if ($net) {
        $s = $s -replace "\s"
    }
    [PSCustomObject]@{
        Length = $s.Length;
        Input = $s;
    } | Write-Output
}

function Invoke-Youtubedl {
    param (
        [string]$url
        ,[switch]$fullMovie
    )
    $target = $url -replace "(^https.+watch\?v=.{11}).+$",'$1'
    $opt = ($fullMovie)? "" : "--extract-audio --audio-format mp3"
    'yt-dlp {0} --restrict-filenames {1} --output "%(title)s_%(id)s.%(ext)s"' -f $target, $opt | Invoke-Expression
}

function ConvertTo-Mp3 {
    param (
        $path
    )
    $file = Get-Item -LiteralPath $path
    $outPath = "{0}.mp3" -f $file.BaseName
    & ffmpeg.exe @("-i", $file.FullName, "-vn", "-qscale:a", 0, $outPath)
}

function Get-Mp3Property {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $sh = New-Object -ComObject Shell.Application
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -eq ".mp3") {
            $nameSpace = $sh.NameSpace($fileObj.Directory.FullName)
            $props = $nameSpace.ParseName($fileObj.Name)
            return [PSCustomObject]@{
                "Name" = $fileObj.Name;
                "FullName" = $fileObj.FullName;
                "Title" = $nameSpace.GetDetailsOf($props, 21)
                "PlayTime" = $nameSpace.GetDetailsOf($props, 27)
            }
        }
    }
    end {}
}
function Test-Url {
    param (
        [parameter(ValueFromPipeline)][string]$inputLine
    )
    begin {
    }
    process {
        try {
            return [PSCustomObject]@{
                "IsValid" = $true;
                "Response" = $(Invoke-WebRequest $inputLine);
            }
        }
        catch {
            return [PSCustomObject]@{
                "IsValid" = $false;
                "Response" = $null;
            }
        }
    }
    end {
    }
}

function Invoke-FileDownload {
    param (
        [parameter(Mandatory)][string]$uri
        ,[string]$name
    )
    if (-not $name) {
        if ("System.Web" -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
            Add-Type -AssemblyName System.Web
        }
        $u = [uri]::new($uri)
        if (-not $u.Authority) {
            return
        }
        $leaf = $u.Segments | Select-Object -Last 1
        $name = [System.Web.HttpUtility]::UrlDecode($leaf)
    }
    $outPath = $pwd.Path | Join-Path -ChildPath $name
    if (Test-Path $outPath) {
        "same file exists! : '{0}'" -f $name | Write-Host -ForegroundColor Magenta
        return
    }
    Invoke-WebRequest -Uri $uri -OutFile $outPath
    "saved as '{0}'" -f $name | Write-Host -ForegroundColor Green

}

function Invoke-Monolith {
    param (
        [parameter(ValueFromPipeline)][string]$inputLine
        ,[string]$outDir
    )
    begin {
        if(-not $outDir) {
            $outDir = "."
        }
        else {
            if (-not (Test-Path $outDir)) {
                New-Item -Path $outDir -ItemType Directory
            }
        }
    }
    process {
        if ($inputLine.Length -gt 0) {
            $childPath = [uri]::new($inputLine).Segments | Select-Object -Skip 1 | ForEach-Object {$_ -replace "/"} | Join-String -Separator "\"
            $outPath = (Resolve-Path $outDir).Path | Join-Path -ChildPath $childPath
            if ($outPath -notmatch "htm(l)?$") {
                $outPath = $outPath | Join-Path -ChildPath "index.html"
            }
            if (Test-Path $outPath){
                "same name '{0}' exists!" -f $outPath | Write-Host -ForegroundColor Magenta
            }
            else {
                $d = $outPath | Split-Path -Parent
                (Resolve-Path $d -ErrorAction SilentlyContinue || New-Item -Path $d -ItemType Directory) > $null
                "archiving as '{0}'..." -f $outPath | Write-Host
                "monolith {0} --no-video --silent --output '{1}'" -f $inputLine, $outPath | Invoke-Expression
                Start-Sleep -Seconds 1
            }
        }
    }
        end {}
}

function akitablogConsole {
    'copy(Array.from(document.querySelectorAll("#calendarplugin-154614 > div.calbody > table > tbody > tr:nth-child(2) > td > table > tbody td a")).map(td=>td.href).join("\r\n"))' | Write-Output
}


function Get-FileProperties {
    param (
        $path
        ,$max = 100
    )
    $file = Get-Item -LiteralPath $path
    if (-not $file) {
        return
    }
    $dirPath = $file.Directory.FullName
    $name = $file.Name
    $shell = New-Object -ComObject Shell.Application
    $shellFolder = $shell.namespace($dirPath)
    $shellFile = $shellFolder.parseName($name)
    0..$max | ForEach-Object {
        [PSCustomObject]@{
            "Id" = $_;
            "Name" = $shellFolder.getDetailsOf($null, $_);
            "Value" = $shellFolder.getDetailsOf($shellFile, $_);
        } | Write-Output
    }
}

if ("PresentationCore" -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
    # required for font search
    Add-Type -AssemblyName PresentationCore
}
function Get-LocalFont {
    $globalFonts = Get-ChildItem "C:\Windows\Fonts" -File
    $userFonts = Get-ChildItem ($env:USERPROFILE | Join-Path -ChildPath "AppData\Local\Microsoft\Windows\Fonts") -File
    $globalFonts + $userFonts | Where-Object {$_.Extension -in @(".ttf", ".otf", ".ttc")} | ForEach-Object {
        $path = $_.Fullname
        try {
            $font = New-Object -TypeName Windows.Media.GlyphTypeface -ArgumentList $path
            return [PSCustomObject]@{
                "Name" = $font.Win32FamilyNames["en-us"];
                "LocalName" = $font.Win32FamilyNames["ja-jp"];
                "Path" = $path;
            }
        }
        catch {
        }
    }
}

function Move-ItemToObsDir {
    $items = @()
    $cbStr = Get-Clipboard
    if ($cbStr.Length) {
        $items = ($cbStr -replace '" "', '"\n"') -split "\n" | ForEach-Object {
            if ($_.StartsWith('"')) {
                return $_ -replace '"'
            }
            return $_ -split " "
        } | Where-Object { Test-Path $_ } | Get-Item
    }
    else {
        $items = @([Windows.Forms.Clipboard]::GetFileDropList() | Get-Item)
    }
    if ($items.Count -lt 1) {
        return
    }
    $p = $null
    foreach ($item in $items) {
        if ($item.Directory) {$p = $item.Directory; break}
        if ($item.Parent) {$p = $item.Parent; break}
    }
    if (-not $p) {
        return
    }
    $dest = $p.Fullname | Join-Path -ChildPath "_obsolete"
    if (-not (Test-Path $dest)) {
        New-Item -Path $dest -ItemType Directory > $null
    }
    "Moving to: " | Write-Host -ForegroundColor Blue -NoNewline
    $dest | Write-Host
    "--------------------------" | Write-Host -ForegroundColor Blue
    $items | ForEach-Object {
        try {
            $_ | Move-Item -Destination $dest -ErrorAction Stop
           "- " | Write-Host -ForegroundColor Blue -NoNewline
           $_.Name | Write-Host
        }
        catch {
            "Same file exists in '{0}'!" -f $dest | Write-Error
        }
    }
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+o" -BriefDescription "move-to-obsDir" -LongDescription "move-to-obsDir" -ScriptBlock {
    [PSBufferState]::new().RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("<#SKIPHISTORY#>Move-ItemToObsDir")
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}
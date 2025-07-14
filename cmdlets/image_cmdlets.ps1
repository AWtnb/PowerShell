
<# ==============================

cmdlets for treating image

                encoding: utf8bom
============================== #>

foreach ($assembly in @("System.Drawing", "System.Windows.Forms")) {
    if ($assembly -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{ $_.GetName().Name })) {
        Add-Type -AssemblyName $assembly
    }
}

class PhotoFile {
    [string]$path
    [string]$name
    [string]$ext
    [datetime]$filler

    PhotoFile([string]$path) {
        $this.path = $path
        $item = Get-Item $path
        $this.name = $item.Name
        $this.ext = $item.Extension
        $this.filler = Get-Date -UnixTimeSeconds 0
    }

    [int] getByteOffset() {
        if (Test-Path $this.path -PathType Container) {
            return -1
        }
        $e = $this.ext.ToLower().Substring(1)
        if ($e -in @("jpeg", "jpg", "webp")) {
            return 0
        }
        if ($e -eq "raf"){
            if ($this.name.StartsWith("_DSF")) {
                return 0x19E
            }
            return 0x17A
        }
        if ($e -eq "cr2") {
            return 0x144
        }
        if ($this.name.StartsWith("MVI_") -and $e -eq "mp4") {
            return 0x160
        }
        return -1
    }

    [datetime] fromExif() {
        $fs = [System.IO.File]::OpenRead($this.path)
        $img = [System.Drawing.Bitmap]::FromStream($fs, $false, $false)
        try {
            $b = $img.GetPropertyItem(0x9003).value
            $bytes =  $b[0..($b.Length - 2)]
            $decoded = [System.Text.Encoding]::ASCII.GetString($bytes)
            return [Datetime]::ParseExact($decoded.Trim(), "yyyy:MM:dd HH:mm:ss", $null)
        }
        catch {
            return $this.filler
        }
        finally {
            $img.Dispose()
            $fs.Close()
        }
    }

    [datetime]getTimestamp() {
        $offset = $this.getByteOffset()
        if ($offset -lt 1) {
            if ($offset -eq 0) {
                return $this.fromExif()
            }
            return $this.filler
        }
        $bytes = Get-Content $this.path -AsByteStream -TotalCount ($offset + 19) | Select-Object -Last 19
        $decoded = [System.Text.Encoding]::ASCII.GetString($bytes)
        return [Datetime]::ParseExact($decoded, "yyyy:MM:dd HH:mm:ss", $null)
    }

    [PSCustomObject] parse([string]$fmt) {
        $ts = $this.getTimestamp().ToString($fmt)
        return [PSCustomObject]@{
            "Name" = $this.name;
            "Timestamp" = $ts;
        }
    }
}

function Rename-PhotoFile {
    param (
        [switch]$execute
    )
    $color = ($execute)? "Cyan" : "White"
    $input | ForEach-Object {
        $p = [PhotoFile]::new($_.FullName).parse("yyyy_MMdd_HHmmss00")
        $newName = "{0}_{1}" -f $p.Timestamp, $p.Name
        try {
            $p.Name | Write-Host -NoNewline
            " => " | Write-Host -ForegroundColor DarkGray -NoNewline
            $newName | Write-Host -ForegroundColor $color
            if ($execute) {
                $_ | Rename-Item -NewName $newName -ErrorAction Stop
            }
        }
        catch {
            "ERROR!: failed to rename '{0}'!" -f $itemName | Write-Host -ForegroundColor Magenta
        }
    }
}

function Get-Mp4Property {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$format = "yyyy_MMdd_HHmmss00"
    )
    begin {
        $sh = New-Object -ComObject Shell.Application
        $filler = Get-Date -Format $format -UnixTimeSeconds 0
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -eq ".mp4") {
            $nameSpace = $sh.NameSpace($fileObj.Directory.FullName)
            $props = $nameSpace.ParseName($fileObj.Name)
            $d = $nameSpace.GetDetailsOf($props, 208) -replace "[\u200e\u200f]", ""
            $ts = ($d.length -lt 1)? $filler : [Datetime]::ParseExact($d, "yyyy/MM/dd H:mm", $null).ToString($format)
            return [PSCustomObject]@{
                "Name" = $fileObj.Name;
                "FullName" = $fileObj.FullName;
                "Timestamp" = $ts
            }
        }
    }
    end {}
}

function Rename-MiteneTimestamp {
    param (
        [switch]$execute
    )
    $color = ($execute)? "Cyan" : "White"
    $format = "yyyy-MM-ddTHHmmss+0900"
    $input | Where-Object Extension -in @(".jpeg", ".jpg", ".webp", ".mp4") | ForEach-Object {
        $itemName = $_.Name
        $timestamp = ($_.Extension -eq ".mp4")? ($_ | Get-Mp4Property -format $format).Timestamp : [PhotoFile]::new($_.FullName).parse($format).Timestamp

        if ($timestamp) {
            $newName = "{0}-{1}" -f $timestamp, $itemName
            try {
                "  '{0}' => '{1}'" -f $itemName, $newName | Write-Host -ForegroundColor $color
                if ($execute) {
                    $_ | Rename-Item -NewName $newName -ErrorAction Stop
                }
            }
            catch {
                "ERROR!: failed to rename '{0}'!" -f $itemName | Write-Error
            }
        }
    }
}

function Invoke-ImageMagickWatermarkFromFile {
    <#
        .EXAMPLE
        ls -file | Invoke-ImageMagickWatermarkFromFile -watermarkPath .\watermark\mark.png
        Invoke-ImageMagickWatermarkFromFile -inputObj .\hogehoge.jpeg -watermarkPath .\watermark\mark.png
    #>
    param(
        [parameter(ValueFromPipeline)]$inputObj,
        [string]$watermarkPath,
        [int]$transparency = 75
    )
    begin {
        $dissolve = 100 - $transparency
        $watermarkPath = (Resolve-Path -Path $watermarkPath).Path
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Fullname -eq $watermarkPath) {
            return
        }
        $fileName = "wm{0}_{1}" -f $transparency, $fileObj.Name
        $savePath = Join-Path -Path $fileObj.Directory.Fullname -ChildPath $fileName
        if (Test-Path $savePath) {
            "'{0}' already exists!" -f $fileName | Write-Error
            return
        }
        $imageWidth = [System.Drawing.Image]::FromFile($fileObj.Fullname).Width
        "magick composite '{0}' -gravity center -resize {1}x -dissolve {2}%x100% '{3}' '{4}'" -f $watermarkPath, $imageWidth, $dissolve, $fileObj.Fullname, $savePath | Invoke-Expression
    }
    end {}
}

function Invoke-ClearWhiteArea {
    param (
        [parameter(ValueFromPipeline)]$inputObj,
        [int]$fuzz = 20
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $outName = $fileObj.basename + "_skeleton" + $fileObj.Extension
        "magick convert '{0}' -fuzz {1}% -transparent white '{2}'" -f $fileObj.Fullname, $fuzz, $outName | Invoke-Expression
    }
    end {}
}

function Invoke-ImageMagickPng2Ico {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -eq ".png") {
            "converting '{0}'..." -f $fileObj.Name | Write-Host
            "magick convert -resize 128x128 {0}.png {0}.ico" -f $fileObj.basename | Invoke-Expression
        }
    }
    end {}
}

function Invoke-ImageMagickShadow {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $fullname = $fileObj.Fullname
        $outName = $fileObj.basename + "_shadow" + $fileObj.extension
        'magick convert "{0}" `( +clone -background white -shadow 100x8+0+0 `) -background none -compose DstOver -flatten -compose Over "{1}"' -f $fullname, $outName | Invoke-Expression
    }
    end {}
}

function Invoke-ImageMagickGrayscale {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $fullname = $fileObj.Fullname
        $outName = "{0}_gray{1}" -f $fileObj.basename, $fileObj.Extension
        'magick convert "{0}" -quality 100 -colorspace Gray {1}' -f $fullname, $outName | Invoke-Expression
    }
    end {}
}

function Invoke-ImageMagickResize {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[int]$width = 256
        ,[int]$height = 0
    )
    begin {
        if ($height -lt 1) {
            $height = $width
        }
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $fullname = $fileObj.Fullname
        $outName = "{0}_{1}{2}" -f $fileObj.basename, $width, $fileObj.Extension
        'magick convert "{0}" -quality 100 -resize {1}x{2} {3}' -f $fullname, $width, $height, $outName | Invoke-Expression
    }
    end {}
}

function Invoke-ImageMagickResizeByWidth {
    param (
        [parameter(ValueFromPipeline)]$inputObj,
        [int]$maxWidth = 500
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $fullname = $fileObj.Fullname
        $outName = "{0}_width{1}{2}" -f $fileObj.basename, $maxWidth, $fileObj.extension
        'magick convert "{0}" -quality 100 -resize {1}x> {2}' -f $fullname, $maxWidth, $outName | Invoke-Expression
    }
    end {}
}

function Convert-ClipboardImage2File {
    param (
        [string]$basename
    )
    $img = [Windows.Forms.Clipboard]::GetImage()
    if(-not $img) {
        return
    }
    $basename = $basename -replace "\.\\(.+)\.png$", '$1'
    if (-not $basename) {
        $basename = Get-Date -Format yyyyMMddHHmmss
    }
    $fullpath = $pwd.ProviderPath | Join-Path -ChildPath ("{0}.png" -f $basename)
    $counter = 0
    while (Test-Path $fullpath) {
        if ($counter -ge 10) {
            Write-Host "failed to save image..." -ForegroundColor Magenta
            return
        }
        $counter += 1
        $fullpath = $pwd.ProviderPath | Join-Path -ChildPath ("{0}-{1}.png" -f $basename, $counter)
        if (-not (Test-Path $fullpath)) {
            break
        }
    }
    $img.save($fullpath)
    "save clipboard image as '{0}'" -f ($fullpath | Split-Path -Leaf) | Write-Host -ForegroundColor Cyan
}
Set-Alias cbImage2file Convert-ClipboardImage2File

Set-PSReadLineKeyHandler -Key "ctrl+I" -BriefDescription "save-clipboard-image" -LongDescription "save-clipboard-image" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("<#SKIPHISTORY#> cbImage2file ")
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}
Set-PSReadLineKeyHandler -Key "ctrl+alt+i" -BriefDescription "save-clipboard-image-with-name" -LongDescription "save-clipboard-image-with-name" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("cbImage2file ")
}

function Invoke-Pngs2Gif {
    param (
        $duration = 400
        ,$outName = "out"
        ,[switch]$noLoop
        ,[switch]$force
    )
    $pngs = $input | Where-Object Extension -eq ".png"
    $outPath = $PWD.ProviderPath | Join-Path -ChildPath "$($outName).gif"
    if ((Test-Path $outPath) -and (-not $force)) {
        Write-Error "Same file exists!"
        return
    }
    $cmd = 'magick -background none -dispose background -loop "{0}" -delay "{1}" {2} "{3}"' -f (($noLoop)? 1 : 0), ($duration / 10), ($pngs | Join-String -DoubleQuote -Separator " "), $outPath
    Invoke-Expression $cmd
}


function Get-ImageSize {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $fs = [System.IO.File]::OpenRead($fileObj.FullName)
        $img = [System.Drawing.Bitmap]::FromStream($fs, $false, $false)
        $info = [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Width" = $img.Width;
            "Height" = $img.Height;
        }
        $img.Dispose()
        $fs.Close()
        return $info
    }
    end {}
}

################################
# publish photo album with moul
################################

function moul {
    & ("C:\Users\{0}\Sync\portable_app\moul\moul.exe" -f $Env:USERNAME) $args
}

function Invoke-ImageMagickWatermarkSignature {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$s
    )
    begin {
        [scriptblock]$tempDirMake = {
            $dirPath = [System.IO.Path]::GetTempPath() | Join-Path -ChildPath $([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dirPath > $null
            return $dirPath
        }

        [scriptblock]$watermark2 = {
            param ([string]$outPath, [string]$s)
            $bmp = [System.Drawing.Bitmap]::new(100, 50)
            $graphic = [System.Drawing.Graphics]::FromImage($bmp)
            $graphic.FillRectangles([System.Drawing.Brushes]::Transparent, $graphic.VisibleClipBounds)
            $fontSize = 12
            $font = [System.Drawing.Font]::new("Meiryo", $fontSize)
            $graphic.DrawString($s, $font, [System.Drawing.Brushes]::White, 10, 10)
            $graphic.Dispose()
            $bmp.Save($outPath)
        }

        [scriptblock]$overlay = {
            param([string]$path, [string]$watermark)
            $file = Get-Item -LiteralPath $path
            $savePath = $file.Directory.Fullname | Join-Path -ChildPath ("wm{0}_{1}" -f $transparency, $file.Name)
            "magick composite '{0}' -gravity SouthEast -quality 100 '{1}' '{2}'" -f $watermark, $file.Fullname, $savePath | Invoke-Expression
        }

        $tmpDirPath = $tempDirMake.InvokeReturnAsIs()

    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -notmatch "jpe?g$") {
            return
        }
        $timestamp = Get-Date -Format yyyyMMddHHmmssff
        $watermarkPath = $tmpDirPath | Join-Path -ChildPath "$($timestamp).png"
        $s = "AWtnb"
        $watermark2.Invoke($watermarkPath, $s)
        $overlay.Invoke($file.FullName, $watermarkPath)
        "Watermarked '{0}' on '{1}'" -f $s, $file.Name | Write-Host -ForegroundColor Cyan
    }
    end {
        Remove-Item -Path $tmpDirPath -Recurse
    }
}

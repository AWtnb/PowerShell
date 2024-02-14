
<# ==============================

cmdlets for treating image

                encoding: utf8bom
============================== #>

foreach ($assembly in @("System.Drawing", "System.Windows.Forms")) {
    if ($assembly -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{ $_.GetName().Name })) {
        Add-Type -AssemblyName $assembly
    }
}


function Get-ExifDate {
    <#
        .SYNOPSIS
        JPEG 画像の EXIF 撮影日時情報を取得する
        .EXAMPLE
        ls | Get-ExifDate
        Get-ExifDate -inputObj .\hogehoge.jpeg
    #>
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        class Jpeg {
            # https://www.84kure.com/blog/2014/07/10/
            static [byte[]] GetExifByteArray ([string]$path) {
                $fs = [System.IO.File]::OpenRead($path)
                $img = [System.Drawing.Bitmap]::FromStream($fs, $false, $false)
                $b = $null
                try {
                    $b = $img.GetPropertyItem(0x9003).value
                }
                catch {}
                finally {
                    $img.Dispose()
                    $fs.Close()
                }
                return $b
            }
        }
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -notmatch "jpe?g$") {
            return
        }
        $byteArray = [Jpeg]::GetExifByteArray($fileObj.FullName)
        if (-not $byteArray) {
            $timestamp = "0000_0000_00000000"
        }
        else {
            $exifTime = [System.Text.Encoding]::ASCII.GetString($byteArray)
            $timestamp = ($exifTime -match "\d")?
                ("{0}_{1}{2}_{3}{4}{5}00" -f ($exifTime -split "[^\d]+")) :
                "0000_0000_00000000"
        }
        return [PSCustomObject]@{
            Name = $fileObj.Name;
            Timestamp = $timestamp;
        }
    }
    end {}
}

function Rename-ExifDate {
    param (
        [switch]$execute
    )
    $color = ($execute)? "Cyan" : "White"
    $input | Where-Object Extension -Match "\.jpe?g$" | ForEach-Object {
        $itemName = $_.Name
        $newName = "{0}_{1}" -f ($_ | Get-ExifDate).Timestamp, $itemName
        try {
            "  '{0}' => '{1}'" -f $itemName, $newName | Write-Host -ForegroundColor $color
            if ($execute) {
                $_ | Rename-Item -NewName $newName -ErrorAction Stop
            }
        }
        catch {
            "ERROR!: failed to rename '{0}'!" -f $itemName | Write-Host -ForegroundColor Magenta
        }
    }
}

function Get-RAFTimestamp {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".RAF") {
            return [PSCustomObject]@{
                Name = $fileObj.Name;
                Timestamp = "";
            }
        }
        $offset = ($fileObj.Name.StartsWith("_DSF")) ? 414 : 378
        $bytes = Get-Content $fileObj.FullName -AsByteStream -TotalCount ($offset + 19) | Select-Object -Last 19
        $decoded = [System.Text.Encoding]::ASCII.GetString($bytes)
        $date = [Datetime]::ParseExact($decoded, "yyyy:MM:dd HH:mm:ss", $null)
        $timestamp = $date.ToString("yyyy_MMdd_HHmmss00")
        return [PSCustomObject]@{
            Name = $fileObj.Name;
            Timestamp = $timestamp;
        }
    }
    end {}
}

function Get-CR2Timestamp {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".CR2") {
            return [PSCustomObject]@{
                Name = $fileObj.Name;
                Timestamp = "";
            }
        }
        $bytes = Get-Content $fileObj.FullName -AsByteStream -TotalCount (324 + 19) | Select-Object -Last 19
        $decoded = [System.Text.Encoding]::ASCII.GetString($bytes)
        $date = [Datetime]::ParseExact($decoded, "yyyy:MM:dd HH:mm:ss", $null)
        $timestamp = $date.ToString("yyyy_MMdd_HHmmss00")
        return [PSCustomObject]@{
            Name = $fileObj.Name;
            Timestamp = $timestamp;
        }
    }
    end {}
}

function Rename-RAFTimestamp {
    param (
        [switch]$execute
    )
    $color = ($execute)? "Cyan" : "White"
    $input | Where-Object Extension -eq ".RAF" | ForEach-Object {
        $itemName = $_.Name
        $timestamp = ($_ | Get-RAFTimestamp).Timestamp
        if ($timestamp) {
            $newName = "{0}_{1}" -f $timestamp, $itemName
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

function Rename-CR2Timestamp {
    param (
        [switch]$execute
    )
    $color = ($execute)? "Cyan" : "White"
    $input | Where-Object Extension -eq ".CR2" | ForEach-Object {
        $itemName = $_.Name
        $timestamp = ($_ | Get-CR2Timestamp).Timestamp
        if ($timestamp) {
            $newName = "{0}_{1}" -f $timestamp, $itemName
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


function xf10Rename {
    param (
        [switch]$execute
    )
    $input | Where-Object {$_.Name.StartsWith("_DSF") -or $_.Name.StartsWith("DSCF")} | ForEach-Object {
        if ($_.Extension -eq ".RAF") {
            $_ | Rename-RAFTimestamp -execute:$execute
        }
        elseif ($_.Extension -eq ".jpg") {
            $_ | Rename-ExifDate -execute:$execute
        }
     }
}

function eosm6Rename {
    param (
        [switch]$execute
    )
    $input | Where-Object Name -Match "IMG" | ForEach-Object {
        if ($_.Extension -eq ".cr2") {
            $_ | Rename-CR2Timestamp -execute:$execute
        }
        elseif ($_.Extension -eq ".jpg") {
            $_ | Rename-ExifDate -execute:$execute
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

function Invoke-ImageMagickResize {
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
    "save clipboard image as '{0}.png'" -f ($fullpath | Split-Path -Leaf) | Write-Host -ForegroundColor Cyan
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
        function _tempDirMake {
            $dirPath = [System.IO.Path]::GetTempPath() | Join-Path -ChildPath $([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dirPath > $null
            return $dirPath
        }

        function _watermark2 ([string]$outPath, [string]$s) {
            $bmp = [System.Drawing.Bitmap]::new(100, 50)
            $graphic = [System.Drawing.Graphics]::FromImage($bmp)
            $graphic.FillRectangles([System.Drawing.Brushes]::Transparent, $graphic.VisibleClipBounds)
            $fontSize = 12
            $font = [System.Drawing.Font]::new("Meiryo", $fontSize)
            $graphic.DrawString($s, $font, [System.Drawing.Brushes]::White, 10, 10)
            $graphic.Dispose()
            $bmp.Save($outPath)
        }

        function _overlay ([string]$path, [string]$watermark) {
            $file = Get-Item -LiteralPath $path
            $savePath = $file.Directory.Fullname | Join-Path -ChildPath ("wm{0}_{1}" -f $transparency, $file.Name)
            "magick composite '{0}' -gravity SouthEast -quality 100 '{1}' '{2}'" -f $watermark, $file.Fullname, $savePath | Invoke-Expression
        }

        $tmpDirPath = _tempDirMake

    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -notmatch "jpe?g$") {
            return
        }
        $timestamp = Get-Date -Format yyyyMMddHHmmssff
        $watermarkPath = $tmpDirPath | Join-Path -ChildPath "$($timestamp).png"
        $s = "AWtnb"
        _watermark2 -outPath $watermarkPath -s $s
        _overlay -path $file.FullName -watermark $watermarkPath
        "Watermarked '{0}' on '{1}'" -f $s, $file.Name | Write-Host -ForegroundColor Cyan
    }
    end {
        Remove-Item -Path $tmpDirPath -Recurse
    }
}

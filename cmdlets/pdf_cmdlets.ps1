﻿
<# ==============================

cmdlets for treating PDF

                encoding: utf8bom
============================== #>

class PyPdf {
    [string]$pyPath
    PyPdf([string]$pyFile) {
        $this.pyPath = $PSScriptRoot | Join-Path -ChildPath "python\pdf" | Join-Path -ChildPath $pyFile
    }
    RunCommand([string[]]$params){
        $fullParams = (@("-B", $this.pyPath) + $params) | ForEach-Object {
            if ($_ -match " ") {
                return ($_ | Join-String -DoubleQuote)
            }
            return $_
        }
        Start-Process -Path python.exe -Wait -ArgumentList $fullParams -NoNewWindow
    }
    static [string[]] getFiles() {
        return Get-ChildItem -File | Where-Object {$_.Extension -eq ".pdf"} | ForEach-Object {".\" + $_.Name} | ForEach-Object {
            if ($_ -match "\s") {
                return $_ | Join-String -DoubleQuote
            }
            return $_
        }
    }
}

function Invoke-PdfConcWithPython {
    param (
        [string]$outName = "conc"
    )

    $outPath = $pwd.ProviderPath | Join-Path -ChildPath "$($outName).pdf"
    if (Test-Path $outPath) {
        "'{0}.pdf' already exists!" -f $outName | Write-Error
        return
    }

    $pdfs = @($input | Where-Object Extension -eq ".pdf")
    if ($pdfs.Count -le 1) {
        return
    }

    $py = [PyPdf]::new("conc.py")

    Use-TempDir {
        $paths = New-Item -Path ".\paths.txt"
        $pdfs.Fullname | Out-File -Encoding utf8NoBOM -FilePath $paths
        $py.RunCommand(@($paths, $outPath))
    }

}
Set-Alias pdfConcPy Invoke-PdfConcWithPython

function Invoke-PdfOverlayWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[string]$overlayPdf
    )
    begin {}
    process {
        $pdfFileObj = Get-Item -LiteralPath $inputObj
        if ($pdfFileObj.Extension -ne ".pdf") {
            Write-Error "Non-pdf file!"
            return
        }
        $overlayPath = (Get-Item -LiteralPath $overlayPdf).FullName
        $py = [PyPdf]::new("overlay.py")
        $py.RunCommand(@($pdfFileObj.FullName, $overlayPath))
    }
    end {}
}

function Invoke-PdfFilenameWatermarkWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[int]$startIdx = 1
        ,[switch]$countThrough
    )
    begin {
        $pdfs = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -eq ".pdf") {
            $pdfs += $file
        }
    }
    end {
        $py = [PyPdf]::new("overlay_filename.py")
        $mode = ($countThrough)? "through" : "single"
        Use-TempDir {
            $in = New-Item -Path ".\in.txt"
            $pdfs.Fullname | Out-File -FilePath $in.FullName -Encoding utf8NoBOM
            $py.RunCommand(@($in.FullName, $startIdx, $mode))
        }
    }
}

function pyGenSearchPdf {
    param (
        [string]$outName = "search_"
    )
    $files = $input | Where-Object {$_.Extension -eq ".pdf"}
    $orgDir = $pwd.ProviderPath
    Use-TempDir {
        $tempDir = $pwd.ProviderPath
        $files | Copy-Item -Destination $tempDir
        Get-ChildItem "*.pdf" | Invoke-PdfUnspreadWithPython
        Get-ChildItem "*.pdf" | Where-Object { $_.BaseName -notmatch "unspread$" } | Remove-Item
        Get-ChildItem "*_unspread.pdf" | Rename-Item -NewName { ($_.BaseName -replace "_unspread$") + $_.Extension }
        Get-ChildItem "*.pdf" | Invoke-PdfFilenameWatermarkWithPython -countThrough
        Get-ChildItem "wm*.pdf" | Invoke-PdfConcWithPython -outName $outName
        Get-ChildItem | Where-Object { $_.BaseName -eq $outName } | Move-Item -Destination $orgDir
    }
}




function Invoke-PdfZipToDiffWithPython {
    param (
        [string]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $oddFile,
        [string]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $evenFile,
        [string]$outName = "outdiff"
    )
    $odd = Get-Item -LiteralPath $oddFile
    $even = Get-Item -LiteralPath $evenFile
    if (($odd.Extension -ne ".pdf") -or ($even.Extension -ne ".pdf")) {
        "non-pdf file!" | Write-Error
        return
    }
    $outPath = $PWD.ProviderPath | Join-Path -ChildPath "$($outName).pdf"
    $py = [PyPdf]::new("zip_to_diff.py")
    $py.RunCommand(@($odd.FullName, $even.FullName, $outPath))
}

function Invoke-PdfExtractWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[int]$from = 1
        ,[int]$to = -1
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("extract.py")
        $py.RunCommand(@($file.Fullname, $from, $to))
    }
    end {}
}
Set-Alias pdfExtractPy Invoke-PdfExtractWithPython

function Invoke-PdfExtractStepWithPython {
    <#
    .EXAMPLE
        Invoke-PdfExtractStepWithPython -path hoge.pdf -froms 1,4,6
    #>
    param (
        [string]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $path
        ,[int[]]$froms
    )
    $file = Get-Item -LiteralPath $path
    if ($file.Extension -ne ".pdf") {
        return
    }
    for ($i = 0; $i -lt $froms.Count; $i++) {
        $f = $froms[$i]
        $t = ($i + 1 -eq $froms.Count)? -1 : $froms[($i + 1)] - 1
        Invoke-PdfExtractWithPython -inputObj $file.FullName -from $f -to $t
    }
}

function Invoke-PdfRotateWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[ValidateSet(90, 180, 270)][int]$clockwise = 180
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("rotate.py")
        $py.RunCommand(@($file.Fullname, $clockwise))
    }
    end {}
}
Set-Alias pdfRotatePy Invoke-PdfRotateWithPython


function Invoke-PdfToImageWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj,
        [int]$dpi = 300
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("to_image.py")
        $py.RunCommand(@($file.FullName, $dpi))
    }
    end {}
}
Set-Alias pdfImagePy Invoke-PdfToImageWithPython

function Invoke-PdfSpreadWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[switch]$singleTopPage
        ,[switch]$toLeft
        ,[switch]$vertical
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("spread.py")
        $params = @($file.FullName)
        if ($singleTopPage) {
            $params += "--singleTopPage"
        }
        if ($toLeft) {
            $params += "--toLeft"
        }
        if ($vertical) {
            $params += "--vertical"
        }
        $py.RunCommand($params)
    }
    end {}
}
Set-Alias pdfSpreadPy Invoke-PdfSpreadWithPython

function pyGenPdfSpreadImg {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[switch]$singleTopPage
        ,[switch]$vertical
        ,[int]$dpi = 300
    )
    $file = Get-Item -LiteralPath $inputObj
    if ($file.Extension -ne ".pdf") {
        return
    }
    Invoke-PdfSpreadWithPython -inputObj $file -singleTopPage:$singleTopPage -vertical:$vertical
    $spreadFilePath = $file.FullName -replace "\.pdf$", "_spread.pdf" | Get-Item
    Invoke-PdfToImageWithPython -inputObj $spreadFilePath -dpi $dpi
}

function Invoke-PdfUnspreadWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[switch]$vertical
        ,[switch]$singleTop
        ,[switch]$singleLast
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("unspread.py")
        $params = @($file.FullName)
        if ($vertical) {
            $params += "--vertical"
        }
        if ($singleTop) {
            $params += "--singleTop"
        }
        if ($singleLast) {
            $params += "--singleLast"
        }
        $py.RunCommand($params)
    }
    end {}
}
Set-Alias pdfUnspreadPy Invoke-PdfUnspreadWithPython


function Invoke-PdfCropCenterWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[ValidateSet("head", "tail", "both")][string]$mode = "both"
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("crop_center.py")
        $py.RunCommand(@($file.Fullname, $mode))
    }
    end {}
}


function Invoke-PdfTrimGalleyMarginWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[float]$tombowPercentH = 8.0
        ,[float]$tombowPercentV = 8.0
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("trim.py")
        $py.RunCommand(@($file.Fullname, $tombowPercentH, $tombowPercentV))
    }
    end {}
}

Set-Alias pdfTrimMarginPy Invoke-PdfTrimGalleyMarginWithPython

function Invoke-PdfSwapWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[string]$newPdf
        ,[int]$swapStartPage = 1
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -eq ".pdf") {
            $py = [PyPdf]::new("swap.py")
            $py.RunCommand(@($file.Fullname, (Get-Item -LiteralPath $newPdf).FullName, $swapStartPage))
        }
    }
    end {}
}

function Invoke-PdfZipPagesWithPython {
    param (
        [string]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $oddFile,
        [string]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $evenFile,
        [string]$outName = "outzip"
    )
    $odd = Get-Item -LiteralPath $oddFile
    $even = Get-Item -LiteralPath $evenFile
    if (($odd.Extension -ne ".pdf") -or ($even.Extension -ne ".pdf")) {
        "non-pdf file!" | Write-Error
        return
    }
    $outPath = $PWD.ProviderPath | Join-Path -ChildPath "$($outName).pdf"
    $py = [PyPdf]::new("zip_pages.py")
    $py.RunCommand(@($odd.Fullname, $even.Fullname, $outPath))
}

function Invoke-PdfUnzipPagesWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
        ,[switch]$evenPages
    )
    begin {
        $opt = ($evenPages)? "--evenPages" : ""
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("unzip_pages.py")
        $py.RunCommand(@($file.FullName, $opt))
    }
    end {}
}

function Invoke-PdfSplitPagesWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
    )
    begin {
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("split.py")
        $py.RunCommand(@($file.FullName))
    }
    end {}
}

function Invoke-PdfTextExtractWithPython {
    param (
        [parameter(ValueFromPipeline)]
        [ArgumentCompleter({[PyPdf]::getFiles()})]
        $inputObj
    )
    begin {
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("get_text.py")
        $py.RunCommand(@($file.FullName))
    }
    end {}
}
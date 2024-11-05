
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
}

function Invoke-GoPdfConc {
    param (
        [string]$outName = "conc"
    )

    $pdfs = @($input | Get-Item | Where-Object Extension -eq ".pdf")
    if ($pdfs.Count -le 1) {
        return
    }
    $gotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\go-pdfconc.exe"
    if (-not (Test-Path $gotool)) {
        "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
        return
    }
    $pdfs.FullName | & $gotool "--outname=$outname"
}
Set-Alias -Name pdfConcGo -Value Invoke-GoPdfConc

function Invoke-DenoPdfConc {
    param (
        [string]$outName = "conc"
    )

    $pdfs = @($input | Get-Item | Where-Object Extension -eq ".pdf")
    if ($pdfs.Count -le 1) {
        return
    }
    $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-conc.exe"
    if (-not (Test-Path $denotool)) {
        "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
        return
    }
    $pdfs.FullName | & $denotool "--outname=$outname"
}
Set-Alias -Name pdfConcDeno -Value Invoke-DenoPdfConc

function Invoke-PdfConcWithPython {
    param (
        [string]$outName = "conc"
    )

    $pdfs = @($input | Get-Item | Where-Object Extension -eq ".pdf")
    if ($pdfs.Count -le 1) {
        return
    }

    $dirs = @($pdfs | ForEach-Object {$_.Directory.Fullname} | Sort-Object -Unique)
    $outDir = ($dirs.Count -gt 1)? $pwd.ProviderPath : $dirs[0]
    $outPath = $outDir | Join-Path -ChildPath "$($outName).pdf"

    if (Test-Path $outPath) {
        "'{0}.pdf' already exists on '{1}'!" -f $outName, $outDir | Write-Error
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
        ,[int]$start = 1
    )
    $files = $input | Where-Object {$_.Extension -eq ".pdf"}
    $orgDir = $pwd.ProviderPath
    Use-TempDir {
        $tempDir = $pwd.ProviderPath
        $files | Copy-Item -Destination $tempDir
        Get-ChildItem "*.pdf" | Invoke-PdfUnspreadWithPython
        Get-ChildItem "*.pdf" | Where-Object { $_.BaseName -notmatch "unspread$" } | Remove-Item
        Get-ChildItem "*_unspread.pdf" | Rename-Item -NewName { ($_.BaseName -replace "_unspread$") + $_.Extension }
        Get-ChildItem "*.pdf" | Invoke-PdfFilenameWatermarkWithPython -countThrough -startIdx $start
        Get-ChildItem "wm*.pdf" | Invoke-PdfConcWithPython -outName $outName
        Get-ChildItem | Where-Object { $_.BaseName -eq $outName } | Move-Item -Destination $orgDir
    }
}




function Invoke-PdfZipToDiffWithPython {
    param (
        [string]
        $oddFile,
        [string]
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
        $inputObj
        ,[int]$from = 1
        ,[int]$to = -1
        ,[string]$outName = ""
    )
    begin {
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -eq ".pdf") {
            $files += $file
        }
    }
    end {
        $files | ForEach-Object {
            $py = [PyPdf]::new("extract.py")
            $py.RunCommand(@($_.Fullname, $from, $to, $outName))
        }
    }
}
Set-Alias pdfExtractPy Invoke-PdfExtractWithPython

function Invoke-PdfExtractStepWithPython {
    <#
    .EXAMPLE
        Invoke-PdfExtractStepWithPython -path hoge.pdf -froms 1,4,6
    #>
    param (
        [string]
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

function Invoke-PdfSpreadWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[switch]$singleTopPage
        ,[switch]$backwards
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
        if ($backwards) {
            $params += "--backwards"
        }
        if ($vertical) {
            $params += "--vertical"
        }
        $py.RunCommand($params)
    }
    end {}
}
Set-Alias pdfSpreadPy Invoke-PdfSpreadWithPython


function Invoke-DenoPdfSpread {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[switch]$vertical
        ,[switch]$opposite
        ,[switch]$asbook
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-spread.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $params = @()
        if ($vertical) {
            $params += "--vertical"
        }
        if ($opposite) {
            $params += "--opposite"
        }
        if ($asbook) {
            $params += "--asbook"
        }
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfSpreadDeno -Value Invoke-DenoPdfSpread

function pyGenPdfSpreadImg {
    param (
        [parameter(ValueFromPipeline)]
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

function Invoke-DenoPdfUnspread {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[switch]$vertical
        ,[switch]$centeredTop
        ,[switch]$centeredLast
        ,[switch]$opposite
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-unspread.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            return
        }
        $files | ForEach-Object {
            $params = @('--path={0}' -f $_.FullName)
            if ($vertical) {
                $params += "--vertical"
            }
            if ($centeredTop) {
                $params += "--centeredTop"
            }
            if ($centeredLast) {
                $params += "--centeredLast"
            }
            if ($opposite) {
                $params += "--opposite"
            }
            & $denotool $params
        }
    }
}
Set-Alias PdfUnspreadDeno Invoke-DenoPdfUnspread

function Invoke-PdfUnspreadWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[switch]$vertical
        ,[switch]$singleTop
        ,[switch]$singleLast
        ,[switch]$backwards
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
        if ($backwards) {
            $params += "--backwards"
        }
        $py.RunCommand($params)
    }
    end {}
}
Set-Alias pdfUnspreadPy Invoke-PdfUnspreadWithPython

function Invoke-PdfTrimGalleyMarginWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[float]$marginHorizontalRatio = 0.08
        ,[float]$marginVerticalRatio = 0.08
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("trim.py")
        $py.RunCommand(@($file.Fullname, $marginHorizontalRatio, $marginVerticalRatio))
    }
    end {}
}

Set-Alias pdfTrimMarginPy Invoke-PdfTrimGalleyMarginWithPython

function Invoke-DenoPdfCropTombow {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-crop.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            return
        }
        $files | ForEach-Object {
            $params = @('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}

Set-Alias PdfCropTombowDeno Invoke-DenoPdfCropTombow


function Invoke-PdfSwapWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[string]$newPdf
        ,[int]$swapStartPage = 1
        ,[int]$swapPageLength = 1
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -eq ".pdf") {
            $py = [PyPdf]::new("swap.py")
            $py.RunCommand(@($file.Fullname, (Get-Item -LiteralPath $newPdf).FullName, $swapStartPage, $swapPageLength))
        }
    }
    end {}
}

function Invoke-PdfInsertWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[string]$newPdf
        ,[int]$insertAfter = 1
        ,[string]$outName = ""
    )
    begin {}
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("insert.py")

        $py.RunCommand(@($file.Fullname, $newPdf, $insertAfter, $outName))
    }
    end {}
}
Set-Alias pdfInsertPy Invoke-PdfInsertWithPython

function Invoke-PdfZipPagesWithPython {
    param (
        [string]
        $oddFile,
        [string]
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
Set-Alias pdfSplitPagesPy Invoke-PdfSplitPagesWithPython

function Invoke-PdfTextExtractWithPython {
    param (
        [parameter(ValueFromPipeline)]
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

function Invoke-DenoPdfExtract {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[int]$from = 1
        ,[int]$to = -1
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-extract.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $params = @("--frompage=$from", "--topage=$to")
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfExtractDeno -Value Invoke-DenoPdfExtract

function Invoke-DenoPdfInsert {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[parameter(Mandatory)][string]$insertPdf
        ,[int]$from = 1
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-insert.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $insertPath = (Get-Item -Path $insertPdf).FullName
        $params = @("--insert=$insertPath", "--frompage=$from")
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfInsertDeno -Value Invoke-DenoPdfInsert

function Invoke-DenoPdfSwap {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[parameter(Mandatory)][string]$embedPdf
        ,[int]$from = 1
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-swap.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $embedPath = (Get-Item -Path $embedPdf).FullName
        $params = @("--embed=$embedPath", "--frompage=$from")
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfSwapDeno -Value Invoke-DenoPdfSwap

function Invoke-DenoPdfWatermark {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$text = ""
        ,[int]$start = 1
        ,[switch]$nombre
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-watermark.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $params = @("--text=$text", "--start=$start")
        if ($nombre) {
            $params += "--nombre"
        }
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfWatermarkDeno -Value Invoke-DenoPdfWatermark

function Invoke-DenoPdfRotate {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[ValidateSet(90, 180, 270)][int]$degree = 90
        ,[int[]]$pages = @()
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-rotate.exe"
        $files = @()
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        if (-not (Test-Path $denotool)) {
            "Not found: {0}" -f $gotool | Write-Host -ForegroundColor Magenta
            return
        }
        $p = $pages -join ","
        $params = @("--degree=$degree", "--pages=$p")
        $files | ForEach-Object {
            $params += ('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}
Set-Alias -Name pdfRotateDeno -Value Invoke-DenoPdfRotate

function Invoke-PdfTitleMetadataModifyWithPython {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[string]$title = "Title"
        ,[switch]$preserveUntouchedData
    )
    begin {
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $py = [PyPdf]::new("modify_metadata.py")
        $py.RunCommand(@($file.FullName, $title, $preserveUntouchedData))
    }
    end {}
}
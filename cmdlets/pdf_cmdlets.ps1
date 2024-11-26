
<# ==============================

cmdlets for treating PDF

                encoding: utf8bom
============================== #>

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

function denoSearchPdf {
    param (
        [string]$outName = "search"
    )
    $files = $input | Where-Object {$_.Extension -eq ".pdf"}
    "Cropping..." | Write-Host -ForegroundColor Blue
    $files | Invoke-DenoPdfCropTombow -ErrorAction Stop
    Get-ChildItem "*.pdf" | Where-Object {$_.BaseName -notmatch "_crop$"} | Remove-Item -ErrorAction Stop
    "Unspreading..." | Write-Host -ForegroundColor Blue
    Get-ChildItem "*_crop.pdf" | Invoke-DenoPdfUnspread -ErrorAction Stop
    Get-ChildItem "*.pdf" | Where-Object {$_.BaseName -notmatch "_crop_unspread$"} | Remove-Item -ErrorAction Stop
    Get-ChildItem "*.pdf" | Rename-Item -NewName {($_.BaseName -replace "_crop_unspread$", "") + ".pdf"} -ErrorAction Stop
    "Watermarking..." | Write-Host -ForegroundColor Blue
    $count = 1
    Get-ChildItem "*.pdf" | ForEach-Object {
        $count += (Invoke-DenoPdfWatermark -inputObj $_ -start $count -nombre -ErrorAction Stop)
    }
    Get-ChildItem "*.pdf" | Where-Object {$_.BaseName -notmatch "_watermark"} | Remove-Item -ErrorAction Stop
    "Concatenating..." | Write-Host -ForegroundColor Blue
    Get-ChildItem "*.pdf" | Sort-Object {$_.Name} | Invoke-DenoPdfConc -outName $outName -ErrorAction Stop
    Get-ChildItem "*.pdf" | Where-Object {$_.BaseName -ne $outName} | Remove-Item -ErrorAction Stop
    "Finished!" | Write-Host -ForegroundColor Blue
}

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

function Invoke-DenoPdfUnspread {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[switch]$vertical
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
            if ($opposite) {
                $params += "--opposite"
            }
            & $denotool $params
        }
    }
}
Set-Alias PdfUnspreadDeno Invoke-DenoPdfUnspread

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

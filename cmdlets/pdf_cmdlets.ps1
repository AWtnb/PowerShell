
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
    try {
        Get-Command go-pdfconc.exe -ErrorAction Stop > $null
    }
    catch {
        "Exe not found" | Write-Host -ForegroundColor Magenta
        $repo = "https://github.com/AWtnb/go-pdfconc"
        "=> Clone and build from {0}" -f $repo | Write-Host
        return
    }
    $pdfs.FullName | & go-pdfconc.exe "--outname=$outname"
}
Set-Alias -Name pdfConcGo -Value Invoke-GoPdfConc

function Get-Pdfjig {
    [OutputType([string])]
    $path = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\repo\pdfjig"
    if (-not (Test-Path $path -PathType Container)) {
        "Not found: {0}" -f $path | Write-Host -ForegroundColor Magenta
        "=> Clone https://github.com/AWtnb/pdfjig" | Write-Host
        return ""
    }
    return $path | Join-Path -ChildPath "main.ts"
}

function Invoke-DenoPdfConc {
    param (
        [string]$outName = "conc"
    )
    $pdfs = @($input | Get-Item | Where-Object Extension -eq ".pdf")
    if ($pdfs.Count -le 1) {
        return
    }
    $jig = Get-Pdfjig
    if ($jig -eq "") {
        $pdfs.FullName | deno run --allow-import --allow-read --allow-write $jig conc -o $outName
    }
}
Set-Alias -Name pdfConcDeno -Value Invoke-DenoPdfConc

function denoSearchPdf {
    param (
        [string]$outName = "search"
    )
    $files = $input | Where-Object {$_.Extension -eq ".pdf"}

    try {
        "Cropping..." | Write-Host
        $files | Invoke-DenoPdfApplyTrimbox
    }
    catch {
        Write-Host "Error: failed to crop files."
        return
    }

    try {
        "Splitting..." | Write-Host
        Get-ChildItem "*_trimbox.pdf" | Invoke-DenoPdfSplit
    }
    catch {
        Write-Host "Error: failed to split pages of files."
        return
    }

    $outDir = "out"
    New-Item -Path $outDir -ItemType Directory -Force > $null
    Get-ChildItem "*_trimbox_split.pdf" | Copy-Item -Destination $outDir

    try {
        Push-Location -Path $outDir
        Get-ChildItem | Rename-Item -NewName {($_.BaseName -replace "_trimbox_split$", "") + ".pdf"}
        "Watermarking..." | Write-Host
        $count = 1
        Get-ChildItem "*.pdf" | ForEach-Object {
            $count += [int](Invoke-DenoPdfWatermark -inputObj $_ -text $_.Name -startNombre $count)
        }
    }
    catch {
        Write-Host "Error: failed to watermark files."
        return
    }
    finally { Pop-Location }

    try {
        Push-Location -Path $outDir
        "Concatenating..." | Write-Host
        Get-ChildItem "*_watermarked.pdf" | Invoke-DenoPdfConc -outName $outName
        Get-Item "$outName.pdf" | Copy-Item -Destination ..
    }
    catch {
        Write-Host "Error: failed to concatenate files."
        return
    }
    finally { Pop-Location }

    "Cleaning..." | Write-Host
    Get-ChildItem -Directory -Filter $outDir | Remove-Item -Recurse
    Get-ChildItem "*_trimbox*.pdf" | Remove-Item

    "Finished!" | Write-Host -ForegroundColor Yellow
}

function Invoke-DenoPdfSpread {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[switch]$vertical
        ,[switch]$opposite
        ,[switch]$singleTopPage
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
        $jig = Get-Pdfjig
        if ($jig -eq "") {
            return
        }
        
        $params = @()
        if ($vertical) {
            $params += "-v"
        }
        if ($opposite) {
            $params += "-o"
        }
        if ($singleTopPage) {
            $params += "-s"
        }
        $files | ForEach-Object {
            deno run --allow-import --allow-read --allow-write $jig spread $params $_.FullName
        }
    }
}
Set-Alias -Name pdfSpreadDeno -Value Invoke-DenoPdfSpread

function Invoke-DenoPdfSplit {
    param (
        [parameter(ValueFromPipeline)]
        $inputObj
        ,[switch]$vertical
        ,[switch]$opposite
    )
    begin {
        $files = @()
        $params = @()
        if ($vertical) {
            $params += "-v"
        }
        if ($opposite) {
            $params += "-o"
        }
    }
    process {
        $file = Get-Item -LiteralPath $inputObj
        if ($file.Extension -ne ".pdf") {
            return
        }
        $files += $file
    }
    end {
        $jig = Get-Pdfjig
        if ($jig -eq "") {
            return
        }
        $files | ForEach-Object {
            deno run --allow-import --allow-read --allow-write $jig split $params $_.FullName
        }
    }
}
Set-Alias PdfSplitDeno Invoke-DenoPdfSplit

function Invoke-DenoPdfApplyTrimbox {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-apply-trimbox.exe"
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
            $repo = "https://github.com/AWtnb/deno-pdf-apply-trimbox"
            "=> Clone and build from {0}" -f $repo | Write-Host
            return
        }
        $files | ForEach-Object {
            $params = @('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}

Set-Alias PdfApplyTrimboxDeno Invoke-DenoPdfApplyTrimbox

function Invoke-DenoPdfTrimMargin {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[double[]]$marginPercentages
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-trim-margin.exe"
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
            $repo = "https://github.com/AWtnb/deno-pdf-trim-margin"
            "=> Clone and build from {0}" -f $repo | Write-Host
            return
        }
        $marginParam = '--margin={0}' -f ($marginPercentages -join ",")
        $files | ForEach-Object {
            $params = @('--path={0}' -f $_.FullName)
            $params += $marginParam
            & $denotool $params
        }
    }
}

Set-Alias PdfTrimMarginDeno Invoke-DenoPdfTrimMargin

function Invoke-DenoPdfUnzipPages {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $denotool = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\deno-pdf-unzip.exe"
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
            $repo = "https://github.com/AWtnb/deno-pdf-unzip"
            "=> Clone and build from {0}" -f $repo | Write-Host
            return
        }
        $files | ForEach-Object {
            $params = @('--path={0}' -f $_.FullName)
            & $denotool $params
        }
    }
}

Set-Alias PdfUnzipPagesDeno Invoke-DenoPdfUnzipPages

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
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            $repo = "https://github.com/AWtnb/deno-pdf-extract"
            "=> Clone and build from {0}" -f $repo | Write-Host
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
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            $repo = "https://github.com/AWtnb/deno-pdf-insert"
            "=> Clone and build from {0}" -f $repo | Write-Host
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
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            $repo = "https://github.com/AWtnb/deno-pdf-swap"
            "=> Clone and build from {0}" -f $repo | Write-Host
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
        ,[string]$startNombre = ""
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
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            $repo = "https://github.com/AWtnb/deno-pdf-watermark"
            "=> Clone and build from {0}" -f $repo | Write-Host
            return
        }
        $params = @("--text=$text", "--startNombre=$startNombre")
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
            "Not found: {0}" -f $denotool | Write-Host -ForegroundColor Magenta
            $repo = "https://github.com/AWtnb/deno-pdf-rotate"
            "=> Clone and build from {0}" -f $repo | Write-Host
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

﻿
<# ==============================

cmdlets for treating Office software

                encoding: utf8bom
============================== #>

class ComController {
    [scriptblock]$clearBlock

    ComController() {
        $this.clearBlock = {
            Get-Variable | Where-Object {$_.Value -is [__ComObject]} | Clear-Variable
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            1 | ForEach-Object {$_} > $null
        }
    }

    Run([scriptblock]$scriptBlock) {
        $scriptBlock.Invoke()
        $this.clearBlock.Invoke()
    }
}


function ConvertTo-WordDocument {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
    )
    begin {
        $files = @()
    }
    process {
        $files += (Get-Item -LiteralPath $inputObj)
    }
    end {
        $cc = [ComController]::new()
        $cc.Run({
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            foreach ($file in $files) {
                $outPath = $file.FullName | Split-Path -Parent | Join-Path -ChildPath ($file.BaseName + ".docx")
                if (Test-Path $outPath) {
                    "ERROR: File already exists -> '{0}'" -f $outPath | Write-Host -ForegroundColor Magenta
                    continue
                }
                $content = $file | Get-Content -Raw
                $doc = $word.Documents.Add()
                $doc.Range(0, 0).Text = $content
                $doc.SaveAs2($outPath)
                $doc.Close($false)
            }
            $word.Quit()
        })
    }
}

function Convert-WordDocument2PDF {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[switch]$acceptRevisions
        ,[switch]$removeComment
    )
    begin {
        $vbConst = [PSCustomObject]@{
            wdExportFormatPDF          = 17;
            wdExportOptimizeForPrint   = 0;
            wdExportAllDocument        = 0;
            wdExportDocumentWithMarkup = 7;
            OpenAfterExport            = $false;
            From                       = $null;
            To                         = $null;
        }
        $docs = @()
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -match "docx?$") {
            $docs += $fileObj
        }
    }
    end {
        $cc = [ComController]::new()
        $cc.Run({
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $docs | ForEach-Object {
                    "converting '{0}' to PDF..." -f $_.Name | Write-Host
                    try {
                        $doc = $word.Documents.Open($_.FullName, $null, $true)
                        if ($removeComment -and $doc.comments.count -gt 0) {
                            $doc.DeleteAllComments()
                            "==> removing all comments..." | Write-Host
                        }
                        if ($acceptRevisions -and $doc.revisions.count -gt 0) {
                            $doc.AcceptAllRevisions()
                            "==> accepting all revisions..." | Write-Host
                        }
                        $outPath = $_.FullName -replace "\.docx?$", ".pdf"
                        $doc.ExportAsFixedFormat($outPath, `
                                $vbConst.wdExportFormatPDF, `
                                $vbConst.OpenAfterExport, `
                                $vbConst.wdExportOptimizeForPrint, `
                                $vbConst.wdExportAllDocument, `
                                $vbConst.From, `
                                $vbConst.To, `
                                $vbConst.wdExportDocumentWithMarkup)
                        $doc.Comments | ForEach-Object {
                            if (-not $_.Scope.Text) {
                                "this comment cannot be displayed on PDF! :`n{0}" -f $_.Range.Text | Write-Host
                            }
                        }
                        $doc.Close($false)
                        "==> finished!" | Write-Host
                    }
                    catch {
                        "ERROR: {0}" -f $_.Exception.Message | Write-Host -ForegroundColor Magenta
                    }
                }
                $word.Quit()
            })
    }
}
Set-Alias word2pdf Convert-WordDocument2PDF


function Convert-Doc2Docx {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $vbConst = [PSCustomObject]@{
            wdFormatXMLDocument = 12
        }
        $docs = @()
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -eq ".doc") {
            $docs += $fileObj
        }
    }
    end {
        $cc = [ComController]::new()
        $cc.Run({
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $docs | ForEach-Object {
                    "saving '{0}' as DOCX..." -f $_.Name | Write-Host
                    try {
                        $doc = $word.Documents.Open($_.FullName, $null, $true)
                        $outPath = $_.FullName -replace "\.doc$", ".docx"
                        $doc.SaveAs2($outPath, $vbConst.wdFormatXMLDocument)
                        $doc.Close($false)
                    }
                    catch {
                        "ERROR: {0}" -f $_.Exception.Message | Write-Error
                    }
                }
                $word.Quit()
            })
    }
}

function Get-OfficeLastSaveTime {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $shell = New-Object -Com Shell.Application
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -in @(".docx", ".xlsx" ,".pptx")) {
            $dir = $shell.NameSpace(($fileObj.Directory -as [string]))
            $parse = $dir.parseName($fileObj.Name)
            $ts = $dir.GetDetailsOf($parse, 154) -replace "[^ \d:/]"
            [PSCustomObject]@{
                "Name"         = $fileObj.Name;
                "LastSaveTime" = (Get-Date $ts);
            } | Write-Output
        }
    }
    end {}
}

function Convert-PowerpointSlide2PDF {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {
        $vbConst = [PSCustomObject]@{
            ppFixedFormatTypePDF          = 2;
            ppFixedFormatIntentPrint      = 2;
            ppPrintHandoutHorizontalFirst = 2;
            ppPrintOutputSlides           = 1;
        }
        $files = @()
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -match "pptx?$") {
            $files += $fileObj
        }
    }
    end {
        $cc = [ComController]::new()
        $cc.Run({
                $powerpoint = New-Object -ComObject PowerPoint.Application
                $powerpoint.Visible = $true
                $files | ForEach-Object {
                    "converting '{0}' to PDF..." -f $_.Name | Write-Host
                    try {
                        $presen = $powerpoint.Presentations.Open($_.FullName, $null, $true)
                        $outPath = $_.FullName -replace "\.pptx$", ".pdf"
                        $max = $presen.Slides.Count
                        $presen.PrintOptions.Ranges.Add(1, $max) > $null
                        $presen.ExportAsFixedFormat(
                            $outPath,
                            $vbConst.ppFixedFormatTypePDF,
                            $vbConst.ppFixedFormatIntentPrint,
                            $false,
                            $vbConst.ppPrintHandoutHorizontalFirst,
                            $vbConst.ppPrintOutputSlides,
                            $false,
                            $presen.PrintOptions.Ranges.Item(1)
                        )
                        $presen.Close()
                        "==> finished!" | Write-Host
                    }
                    catch {
                        "ERROR: {0}" -f $_.Exception.Message | Write-Host -ForegroundColor Magenta
                    }
                }
                $powerpoint.Quit()
            })
    }
}

function Join-PowerpointSlides {
    param (
        [string]$outName = "conc"
        ,[switch]$force
    )
    $slides = $input | Where-Object {$_.Extension -eq ".pptx"}
    if ($slides.Count -lt 1) {
        return
    }
    $first = $slides | Select-Object -First 1
    $outPath = $first.FullName | Split-Path -Parent | Join-Path -ChildPath "$outName.pptx"
    if (Test-Path $outPath) {
        if (-not $force) {
            "'{0}' already exists." -f $outPath | Write-Error
            return
        }
    }
    $first | Copy-Item -Destination $outPath
    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpoint.Visible = $true
    $presen = $powerpoint.Presentations.Open($outPath)
    $slides | Select-Object -Skip 1 | ForEach-Object {
        $idx = $presen.Slides.Count
        $presen.Slides.InsertFromFile($_.FullName, $idx) > $null
    }

}

<# ==============================

cmdlets for treating Office software

                encoding: utf8bom
============================== #>


function Clear-ComReference {
    Get-Variable | Where-Object {$_.Value -is [__ComObject]} | Clear-Variable -Verbose
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Convert-WordDocument2PDF {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
    )
    begin {
        $vbConst = [PSCustomObject]@{
            wdExportFormatPDF = 17;
            wdExportOptimizeForPrint = 0;
            wdExportAllDocument = 0;
            wdExportDocumentWithMarkup = 7;
            OpenAfterExport = $false;
            From = $null;
            To = $null;
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
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $docs | ForEach-Object {
            "converting '{0}' to PDF..." -f $_.Name | Write-Host
            try {
                $doc = $word.Documents.Open($_.FullName, $null, $true)
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
        Clear-ComReference
    }
}
Set-Alias word2pdf Convert-WordDocument2PDF


function Convert-Doc2Docx {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
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
        Clear-ComReference
    }
}

function Get-OfficeLastSaveTime {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
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
                "Name" = $fileObj.Name;
                "LastSaveTime" = (Get-Date $ts);
            } | Write-Output
        }
    }
    end {}
}

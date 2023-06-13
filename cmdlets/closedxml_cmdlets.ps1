
<# ==============================

cmdlets for treating Microsoft Excel with ClosedXML

                encoding: utf8bom
============================== #>

<#
ClosedXml Dependencies (2021-06-08)
https://www.nuget.org/packages/ClosedXML

.NETFramework 4.0
    DocumentFormat.OpenXml (>= 2.7.2)
    ExcelNumberFormat (>= 1.0.10)
.NETFramework 4.6
    DocumentFormat.OpenXml (>= 2.7.2)
    ExcelNumberFormat (>= 1.0.10)
    Microsoft.CSharp (>= 4.7.0)
.NETStandard 2.0
    DocumentFormat.OpenXml (>= 2.7.2)
    ExcelNumberFormat (>= 1.0.10)
    Microsoft.CSharp (>= 4.7.0)
    System.Drawing.Common (>= 4.5.0)

    runtime: netstandart2.0

#>

$PSScriptRoot | Join-Path -ChildPath "lib\closedxml" | Get-ChildItem | Where-Object {$_.Extension -eq ".dll"} | ForEach-Object {
    if ($_.Basename -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
        Add-Type -Path $_.FullName
    }
}

Class XLSX {
    $workBook
    $name

    XLSX([string]$path) {
        $file = Get-Item -LiteralPath $path
        if ($file.Extension -ne ".xlsx") {
            return
        }
        $this.name = $file.Name
        $this.workBook = New-Object ClosedXML.Excel.XLWorkbook($file.FullName)
    }

    [PSCustomObject[]] GetData() {
        if (-not $this.workBook) {
            return $null
        }
        return $this.workBook.WorkSheets | ForEach-Object {
            return [PSCustomObject]@{
                "Name" = $_.Name;
                "Cells" = ($_.CellsUsed() | ForEach-Object {
                    return [PSCustomObject]@{
                        "Column" = $_.Address.ColumnNumber;
                        "ColumnLetter" = $_.Address.ColumnLetter;
                        "Row" = $_.Address.RowNumber;
                        "Address" = ($_.Address.ColumnLetter + $_.Address.RowNumber);
                        "Value" = $_.Value;
                    }
                });
            }
        }
    }

}

Class ParseExcel {

    static [PSCustomObject] GetData ([PSCustomObject]$sheet) {
        return @($sheet.CellsUsed() | ForEach-Object {
            return [PSCustomObject]@{
                "Column" = $_.Address.ColumnNumber;
                "ColumnLetter" = $_.Address.ColumnLetter;
                "Row" = $_.Address.RowNumber;
                "Address" = ($_.Address.ColumnLetter + $_.Address.RowNumber);
                "Value" = $_.Value;
            }
        })
    }

    static [PSCustomObject] GetAllSheets ([PSCustomObject]$book) {
        return $book.Worksheets | ForEach-Object {
            return [PSCustomObject]@{
                "Name" = $_.Name;
                "Cells" = [ParseExcel]::GetData($_);
            }
        }
    }

}


function Invoke-ClosedXmlExcelParse {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $x = [XLSX]::new($fileObj.FullName)
        return [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Data" = $x.GetData()
        }
    }
    end {}
}

function Invoke-ClosedXmlExcelSearch {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$pattern
        ,[switch]$case
    )
    begin {
        $reg = ($case)? [regex]::new($pattern) : [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $x = [XLSX]::new($fileObj.FullName)
        $x.GetData() | ForEach-Object {
            $sheet = $_
            foreach ($c in $sheet.Cells) {
                if ($reg.IsMatch($c.Value)) {
                    [PSCustomObject]@{
                        "Value" = $c.Value;
                        "Sheet" = $sheet.Name;
                        "Address" = $c.Address;
                        "FileInfo" = $fileObj;
                    } | Write-Output
                }
            }
        }
    }
    end {}
}



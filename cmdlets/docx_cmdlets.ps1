﻿
<# ==============================

cmdlets for treating word without openning file

                encoding: utf8bom
============================== #>

if ("System.IO.Compression.Filesystem" -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
    Add-Type -AssemblyName System.IO.Compression.Filesystem
}

class Docx2 {

    [string]$Path
    [string]$Status
    [string]$DocumentXml
    [string]$CommentXml

    Docx2([string]$path) {
        $this.Path = $path
        if ($this.IsOpened()) {
            $this.Status = "FILEOPENED"
        }
        else {
            $this.Status = "OK"
            $rawXml = $this.ReadData("word/document.xml")
            $this.DocumentXml = $rawXml.Replace("<w:t xml:space=`"preserve`">", "<w:t>") -replace "<mc:FallBack>.+?</mc:FallBack>"
            $this.CommentXml = $this.ReadData("word/comments.xml")
        }
    }

    [bool] IsOpened() {
        $stream = $null
        $inAccessible = $false
        try {
            $stream = [System.IO.FileStream]::new($this.Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch {
            $inAccessible = $true
        }
        finally {
            if($stream) {
                $stream.Close()
            }
        }
        return $inAccessible
    }

    [string] ReadData([string]$relPath) {
        $content = ""
        $archive = [IO.Compression.Zipfile]::OpenRead($this.Path)
        $entry = $archive.GetEntry($relPath)
        if ($entry) {
            $stream = $entry.Open()
            $reader = New-Object IO.StreamReader($stream)
            $content = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
        }
        $archive.Dispose()
        return $content
    }

    static [string] GetNodeText([string]$node) {
        $m = [regex]::Matches($node, "(?<=<w:t>).+?(?=</w:t>)")
        return ($m.Value -replace "&amp;", "&") -join ""
    }

    [string[]] GetParagraphNodes() {
        $m = [regex]::Matches($this.DocumentXml, "<w:p [^>]+?>.+?</w:p>")
        return $m.Value
    }

    [string[]] GetParagraphs() {
        return $this.GetParagraphNodes().ForEach({[Docx2]::GetNodeText($_)})
    }

    [string[]] GetRangeNodes() {
        $m = [regex]::Matches($this.DocumentXml, "(<w:r>.+?</w:r>)|(<w:r w:[^>]+?>.+?</w:r>)")
        return $m.Value
    }

    [string[]] FilterRange([regex]$pattern) {
        $arr = New-Object System.Collections.ArrayList
        $sb = New-Object System.Text.StringBuilder
        $nodes = $this.GetRangeNodes()
        foreach ($n in $nodes) {
            if ($pattern.IsMatch($n)) {
                $sb.Append($n) > $null
            }
            else {
                $s = $sb.ToString()
                if ($sb.Length) {
                    $arr.Add($s) > $null
                }
                $sb.Clear()
            }
        }
        return $arr
    }

    [string[]] GetCommentNodes() {
        $m = [regex]::Matches($this.CommentXml, "<w:comment w:[^>]*?>.*?</w:comment>")
        return $m.Value
    }

}


function Get-DocxContent {
    <#
        .EXAMPLE
        Get-DocxContent .\test.docx
        .EXAMPLE
        ls | Get-DocxContent
    #>
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $fullPath = $fileObj.FullName
        $docx = [Docx2]::new($fullPath)
        return [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Status" = $docx.Status;
            "Paragraphs" = $docx.GetParagraphs();
        }
    }
    end {}
}

function Get-DocxParagraph {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        $d = Get-DocxContent $fileObj
        if ($d.Status -eq "OK") {
            $d.Paragraphs
        }
    }
    end {}
}

function Get-DocxMarkeredString {
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$color = "yellow"
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $fullPath = $fileObj.FullName
        $docx = [Docx2]::new($fullPath)
        return [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Status" = $docx.Status;
            "Decorated" = $docx.FilterRange("<w:highlight w:val=`"$($color)`"/>").ForEach({ [Docx2]::GetNodeText($_) });
        }
    }
    end {}
}

function Get-DocxBoldString {
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $fullPath = $fileObj.FullName
        $docx = [Docx2]::new($fullPath)
        return [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Status" = $docx.Status;
            "Decorated" = $docx.FilterRange("<w:b/>").ForEach({ [Docx2]::GetNodeText($_) });
        }
    }
    end {}
}

function Invoke-DocxGrep {
    <#
        .EXAMPLE
        ls | Invoke-DocxGrep -pattern "ほげ"
        Invoke-DocxGrep -inputObj .\hoge.docx "ほげ"
    #>
    param (
        [parameter(Mandatory, ValueFromPipeline, ValueFromRemainingArguments)]$inputObj
        ,[parameter(Mandatory)][string]$pattern
        ,[switch]$case
        ,[switch]$asObject
    )
    begin {
        $result = New-Object System.Collections.ArrayList
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj[0]
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $docx = [Docx2]::new($fileObj.FullName)
        if ($docx.Status -eq "OK") {
            $found = $docx.GetParagraphs() | Select-String -Pattern $pattern -AllMatches -CaseSensitive:$case | ForEach-Object {
                return [PSCustomObject]@{
                    "Line" = $_.Line;
                    "Matches" = $_.Matches.Value;
                }
            }
        }
        else {
            $found = @()
        }
        $result.Add([PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Path" = $fileObj.FullName;
            "Status" = $docx.Status;
            "Found" = $found;
        }) > $null
    }
    end {
        if ($asObject) {
            return $result
        }
        $result | ForEach-Object {
            if($_.Status -ne "OK") {
                "{0}:{1}" -f ($_.Path | Resolve-Path -Relative), $_.Status | Write-Host -ForegroundColor Magenta
                return
            }
            if ($_.Found.Count) {
                "{0}:" -f ($_.Path | Resolve-Path -Relative) | Write-Host -ForegroundColor Blue
                $_.Found | ForEach-Object {
                    "§" | Write-Host -ForegroundColor Blue -NoNewline
                    $_.Line | hilight -Pattern $pattern -case:$case -backgroundcolor "Yellow"
                }
                "========== {0} matched in '{1}' ==========`n" -f $_.Found.Matches.Count, $_.Name | Write-Host -ForegroundColor Green
            }
        }
    }
}

function Get-DocxMatchPattern {
    <#
        .EXAMPLE
        ls *docx | Get-DocxMatchPattern -pattern ".+?さん"
        Get-DocxMatchPattern -inputObj .\hoge.docx -pattern ".+?さん"
    #>
    param (
        [parameter(ValueFromPipeline)]$inputObj
        ,[string]$pattern
    )
    begin {
        $result = New-Object System.Collections.ArrayList
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $docx = [Docx2]::new($fileObj.FullName)
        if ($docx.Status -ne "OK") {
            "'{0}' is opend by other process!" -f $fileObj.Name | Write-Host -ForegroundColor Magenta
            return
        }

        $grep = $docx.GetParagraphs() | Select-String -Pattern $pattern -AllMatches -CaseSensitive:$case
        $grep.Matches.Value | Group-Object | ForEach-Object {
            $result.Add(
                [PSCustomObject]@{
                    "Match" = $_.Name;
                    "Count" = $_.Count;
                    "File" = $fileObj;
                }
            ) > $null
        }
    }
    end {
        return $result
    }
}

function Get-DocxComment {
    <#
        .EXAMPLE
        Get-DocxComment ./hoge.docx
        .EXAMPLE
        ls | Get-DocxComment
    #>
    param (
        [parameter(ValueFromPipeline)]$inputObj
    )
    begin {}
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $docx = [Docx2]::new($fileObj.FullName)
        $comments = @()
        if($docx.Status -eq "OK") {
            $comments = $docx.GetCommentNodes().ForEach({
                return [PSCustomObject]@{
                    "Author" = [regex]::Match($_, "(?<=w:author=).+?(?= w.date)").value;
                    "Text" = $_ -replace "<[^>]+?>";
                }
            })
        }
        return [PSCustomObject]@{
            "Name" = $fileObj.Name;
            "Status" = $docx.Status;
            "Comments" = $comments;
        }
    }
    end {}
}

function Invoke-DocxCommentGrep {
    param (
        [parameter(Mandatory, ValueFromPipeline, ValueFromRemainingArguments)]$inputObj
        ,[parameter(Mandatory)][string]$pattern
        ,[switch]$case
        ,[switch]$asObject
    )
    begin {
        $result = New-Object System.Collections.ArrayList
        $reg = ($case)? [regex]::new($pattern) : [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj[0]
        if ($fileObj.Extension -ne ".docx") {
            return
        }
        $docx = [Docx2]::new($fileObj.FullName)
        if ($docx.Status -ne "OK") {
            "'{0}' is opened by other process!" -f $fileObj.Name | Write-Error
            return
        }
        $docx.GetCommentNodes() | ForEach-Object {
            $text = $_ -replace "<[^>]+?>";
            $m = $reg.Matches($text)
            if ($m.Count) {
                $author = [regex]::Match($_, "(?<=w:author=).+?(?= w.date)").value;
                $result.Add([PSCustomObject]@{
                    "Name" = $fileObj.Name;
                    "Path" = $fileObj.FullName;
                    "Line" = $text;
                    "Author" = $author;
                    "MatcheValues" = $m.Value;
                }) > $null
            }
        }
    }
    end {
        if ($asObject) {
            return $result
        }
        $result | Group-Object -Property Path | ForEach-Object {
            $path = $_.Name
            "{0}:" -f ($path | Resolve-Path -Relative) | Write-Host -ForegroundColor Blue -NoNewline
            $_.Group.MatcheValues.Count | Write-Host -ForegroundColor Green
            $_.Group | ForEach-Object {
                "By {0}:" -f $_.Author | Write-Host -ForegroundColor DarkGreen -NoNewline
                $_.Line | hilight -pattern $pattern -case:$case
            }
        }
    }
}

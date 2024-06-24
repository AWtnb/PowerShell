
<# ==============================

cmdlets for treating markdown

                encoding: utf8bom
============================== #>

function Invoke-MarkdownRenderPython {
    param (
        [parameter(Mandatory)]
        [ArgumentCompleter({
            return (Get-ChildItem "*.md").Name | ForEach-Object {".\" + $_} | ForEach-Object {
                if ($_ -match "\s") {
                    return $_ | Join-String -DoubleQuote
                }
                return $_
            }
        })][string]$path
        ,[switch]$invoke
        ,[switch]$noDefaultCss
        ,[switch]$openDir
        ,[string]$faviconUnicode = "1F4DD"
    )

    if ($path -match "[\r\n]") {
        return
    }
    if (-not (Test-Path $path)) {
        return
    }
    $pyCodePath = $PSScriptRoot | Join-Path -ChildPath "python\markdown\md.py"
    $mdPath = (Get-Item -LiteralPath $path).FullName
    $params = @(
        "-B",
        $pyCodePath,
        $mdPath
    ) | ForEach-Object {
        if ($_ -match "\s") {
            return $_ | Join-String -DoubleQuote
        }
        return $_
    }
    $params += "--faviconUnicode $faviconUnicode"
    if ($noDefaultCss) {
        $params += "--noDefaultCss"
    }
    if ($invoke) {
        $params += "--invoke"
    }
    Start-Process -path python.exe -wait -NoNewWindow -ArgumentList $params

    if ($openDir) {
        $path | Split-Path -Parent | Invoke-Item
    }
}
Set-Alias mdRenderPy Invoke-MarkdownRenderPython

Set-PSReadLineKeyHandler -Key "ctrl+M" -BriefDescription "render-as-markdown" -LongDescription "Render-as-markdown" -ScriptBlock {
    $cbFile = [Windows.Forms.Clipboard]::GetFileDropList() | Get-Item
    $path = ($cbFile)? $cbFile.FullName : (Get-Clipboard | Select-Object -First 1).Replace('"', "")
    mdRenderPy $path -invoke
    [Microsoft.PowerShell.PSConsoleReadLine]::AddToHistory("mdRenderPy $path -invoke")
}


function mdFrontmatter {
@"
---
title:
styles:
  - h1:
      font-size: 2rem
  - blockquote > p:
      text-indent: -2em
      margin-left: 2.5em
      line-height: 1.5
  - blockquote p:
      text-wrap: pretty
      word-break: break-all
  - blockquote > p > em:
      background-color: "#d0ee23"
css-vars:
  # width-container: 50rem
---
"@ | Write-Output
}
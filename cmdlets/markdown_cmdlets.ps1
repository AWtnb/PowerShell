
<# ==============================

cmdlets for treating markdown

                encoding: utf8bom
============================== #>

function mdLess {
    param (
        [switch]$edit
        ,[switch]$gist
    )
    $lessPath = $PSScriptRoot | Join-Path -ChildPath "python\markdown\markdown.less"
    if ($edit) {
        "code $lessPath" | Invoke-Expression
        return
    }
    $lessContent = Get-Item -LiteralPath $lessPath | Get-Content
    if ($gist) {
        Start-Process "https://gist.github.com/AWtnb/46d2b7a4cc53fb1a1883ff3fb1777477/edit"
        $lessContent | Set-Clipboard
        Write-Host "copied source of 'markdown.less'!" -ForegroundColor Green
        return
    }
    return $lessContent
}

Set-PSReadLineKeyHandler -Key "alt+M" -BriefDescription "markdown-less" -LongDescription "markdown-less" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("<#SKIPHISTORY#> mdLess -")
    [Microsoft.PowerShell.PSConsoleReadLine]::MenuComplete()
}



function Invoke-MarkdownRenderPython {
    param (
        [string]$path
        ,[switch]$invoke
        ,[switch]$noDefaultCss
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

    if ($invoke) {
        Hide-ConsoleWindow
    }

}
Set-Alias mdRenderPy Invoke-MarkdownRenderPython

Set-PSReadLineKeyHandler -Key "ctrl+M" -BriefDescription "render-as-markdown" -LongDescription "Render-as-markdown" -ScriptBlock {
    $cbFile = [Windows.Forms.Clipboard]::GetFileDropList() | Get-Item
    $path = ($cbFile)? $cbFile.FullName : (Get-Clipboard | Select-Object -First 1).Replace('"', "")
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("mdRenderPy '$path' -invoke")
    [Microsoft.PowerShell.PSConsoleReadLine]::AcceptLine()
}

function Get-FaviconMarkup {
    param (
        [parameter(Mandatory)][string]$codepoint
    )
    if ("System.Web" -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
        Add-Type -AssemblyName System.Web
    }
    $svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text x="50%" y="50%" style="dominant-baseline:central;text-anchor:middle;font-size:90px;">&#x{0};</text></svg>' -f $codepoint
    $encoded = [regex]::Replace($svg, "[^a-z/\.]", { [System.Web.HttpUtility]::UrlEncode($args.Value) })
    return ('<link rel="icon" href="data:image/svg+xml,{0}">' -f $encoded.Replace("+", "%20"))
}

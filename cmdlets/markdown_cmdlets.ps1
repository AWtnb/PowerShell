
<# ==============================

cmdlets for treating markdown

                encoding: utf8bom
============================== #>

function Invoke-Markdown2Html {
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
        ,[switch]$plain
        ,[switch]$invoke
    )

    if ($path -match "[\r\n]") {
        return
    }
    if (-not (Test-Path $path)) {
        return
    }

    $exe = $env:USERPROFILE | Join-Path -ChildPath "Personal\tools\bin\m2h.exe"
    if (-not (Test-Path $exe)) {
        "Not found: {0}" -f $exe | Write-Host -ForegroundColor Magenta
        $repo = "https://github.com/AWtnb/m2h"
        "=> Clone and build from {0}" -f $repo | Write-Host
        return
    }

    $md = Get-Item -LiteralPath $path
    $suf = Get-Date -Format "_yyyyMMdd"
    $params = @(
        "-src",
        $md.FullName,
        "-suffix",
        $suf
    )
    if ($plain) {
        $params += "-plain"
    }

    Start-Process -path $exe -wait -NoNewWindow -ArgumentList $params

    if ($invoke) {
        $outPath = $md.Directory.FullName | Join-Path -ChildPath ($md.BaseName + $suf + ".html")
        Invoke-Item $outPath
    }

}
Set-Alias m2hgo Invoke-Markdown2Html

Set-PSReadLineKeyHandler -Key "ctrl+M" -BriefDescription "render-as-markdown" -LongDescription "Render-as-markdown" -ScriptBlock {
    $cbFile = [Windows.Forms.Clipboard]::GetFileDropList() | Get-Item
    $path = ($cbFile)? $cbFile.FullName : (Get-Clipboard | Select-Object -First 1).Replace('"', "")
    m2hgo $path -invoke
    [Microsoft.PowerShell.PSConsoleReadLine]::AddToHistory("m2hgo $path -invoke")
}


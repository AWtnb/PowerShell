
<# ==============================

cmdlets for treating markdown

                encoding: utf8bom
============================== #>

function Invoke-Markdown2Html {
    param (
        [parameter(Mandatory)]
        [string]$path
        ,[switch]$plain
        ,[switch]$invoke
    )

    $path = $path.Trim()
    if ($path -match "[\r\n]") {
        return
    }
    if (-not (Test-Path $path)) {
        return
    }

    try {
        Get-Command m2h.exe -ErrorAction Stop > $null
    }
    catch {
        "Exe not found" | Write-Host -ForegroundColor Magenta
        $repo = "https://github.com/AWtnb/m2h"
        "=> Clone and build from {0}" -f $repo | Write-Host
        return
    }

    $md = Get-Item -LiteralPath $path
    $suf = Get-Date -Format "_yyyyMMdd"
    $params = @(
        ("--src={0}" -f $md.FullName),
        ("--suffix={0}" -f $suf)
    )
    if ($plain) {
        $params += "--plain"
    }

    m2h.exe $params

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


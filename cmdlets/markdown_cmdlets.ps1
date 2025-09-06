
<# ==============================

cmdlets for treating markdown

                encoding: utf8bom
============================== #>

function Invoke-MarkdownDocumentServer {
    param (
        [parameter(Mandatory)]
        [string]$path
        ,[switch]$plain
        ,[switch]$export
    )

    $path = $path.Trim()
    if ($path -match "[\r\n]") {
        return
    }
    if (-not (Test-Path $path)) {
        return
    }

    try {
        Get-Command ddserv.exe -ErrorAction Stop > $null
    }
    catch {
        "Exe not found" | Write-Host -ForegroundColor Magenta
        $repo = "https://github.com/AWtnb/ddserv"
        "=> Clone and build from {0}" -f $repo | Write-Host
        return
    }

    $md = Get-Item -LiteralPath $path
    $params = @(
        ("--src={0}" -f $md.FullName)
    )
    if ($plain) {
        $params += "--plain"
    }
    if ($export) {
        $params += "--export"
    }
    else {
        Start-Process "http://localhost:8080"
    }

    ddserv.exe $params

}



<# ==============================

Web search

                encoding: utf8bom
============================== #>

if ("System.Web" -notin ([System.AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object{$_.GetName().Name})) {
    Add-Type -AssemblyName System.Web
}

function Invoke-GoogleSearch {
    [CmdletBinding(DefaultParameterSetName="global")]
    param (
        [Parameter(ParameterSetName="global")]
            [Switch]$global
        ,[Parameter(ParameterSetName="image")]
            [Switch]$image
        ,[Parameter(ParameterSetName="map")]
            [Switch]$map
        ,[Parameter(ParameterSetName="scholar")]
            [Switch]$scholar
        ,[Parameter(ParameterSetName="yu")]
            [Switch]$yu
        ,[switch]$strict
        ,[Parameter(ValueFromRemainingArguments)]
            [string[]]$s
    )
    $keyword = $s | Join-String -Separator " " -DoubleQuote:$strict

    $url = switch ($PsCmdlet.ParameterSetName) {
        "global" {"http://www.google.co.jp/search?q={0}"; break}
        "image" {"https://www.google.com/search?tbm=isch&q={0}"; break}
        "map" {"https://www.google.co.jp/maps/search/{0}"; break}
        "scholar" {"https://scholar.google.co.jp/scholar?q={0}"; break}
        "yu" {"http://www.google.co.jp/search?tbs=li:1&q=site%3Ayuhikaku.co.jp%20intitle%3A{0}"; break}
    }
    Start-Process ([string]$url -f [System.Web.HttpUtility]::UrlEncode($keyword))
}
Set-Alias google Invoke-GoogleSearch

Set-PSReadLineKeyHandler -Key "ctrl+q,spacebar","ctrl+q,i", "ctrl+q,m", "ctrl+q,s", "ctrl+q,y" -ScriptBlock {
    param($key, $arg)

    $opt = switch ($key.KeyChar) {
        <#case#> " " {''; break}
        <#case#> "i" {'-image '; break}
        <#case#> "m" {'-map '; break}
        <#case#> "s" {'-scholar '; break}
        <#case#> "y" {'-yu '; break}
    }
    $command = 'google {0}' -f $opt
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert($command)
}

function twitterSearch ([switch]$ja, [switch]$strict) {
    $keyword = $args -join " "
    if ($strict) {
        $keyword = '"' + ($keyword -replace " ", '" "') + '"'
    }
    if ($ja) {
        $keyword += " lang:ja"
    }
    $encoded = [System.Web.HttpUtility]::UrlEncode($keyword)
    Start-Process ("https://nitter.net/search?f=tweets&q={0}" -f $encoded)
    Hide-ConsoleWindow
}
Set-Alias tw twitterSearch

function amazonPhotoJump {
    param (
        [int]$y,
        [int]$m,
        [ValidateSet("parallel", "mix", "jpg", "jpeg")][string]$mode = "parallel"
    )
    if (-not $y) {
        $y = (Get-Date).Year
    }
    if ($mode -eq "parallel") {
        @("jpeg", "jpg") | ForEach-Object {
            Start-Process ("https://www.amazon.co.jp/photos/search/all/{0}?lcf=time&timeYear={1}&timeMonth={2}" -f $_, $y, $m)
        }
        return
    }
    $param = ($mode -eq "mix")?
        ("all?lcf=time&timeYear={0}&timeMonth={1}" -f $y, $m) :
        ("search/all/{0}?lcf=time&timeYear={1}&timeMonth={2}" -f $mode, $y, $m)
    Start-Process ("https://www.amazon.co.jp/photos/{0}" -f $param)
}

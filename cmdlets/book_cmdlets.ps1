
<# ==============================

cmdlets for treating book information

                encoding: utf8bom
============================== #>

class YBookCode {

    [string]$CheckDigit
    [string]$FullCode

    YBookCode([string]$five) {
        $pad = $five.PadLeft(5, "0")
        $this.CheckDigit = [YBookCode]::GetCheckDigit($pad)
        $this.FullCode = "9784641" + $pad + $this.CheckDigit
    }

    [string] Format() {
        return $this.FullCode -replace "(\d{3})(\d)(\d{3})(\d{5})(\d)", '$1-$2-$3-$4-$5'
    }

    [string] ToUrl() {
        return "http://www.yuhikaku.co.jp/books/detail/" + $this.FullCode
    }

    [void] Run() {
        Start-Process $this.ToUrl()
    }

    static [string] GetCheckDigit([string]$five) {
        $code12 =  "9784641" + $five
        $sum = 0
        1..12 | ForEach-Object {
            $n = [char]::GetNumericValue($code12[$_ - 1])
            $sum += (($_ % 2) -ne 0)? $n : $n * 3
        }
        $cd = (10 - ($sum % 10)) % 10
        return $($cd -as [string])
    }

}


function Resolve-YBookCode {
    <#
        .EXAMPLE
        Resolve-YBookCode 12345
    #>
    param (
        [string]$code
        ,[switch]$run
    )
    $yc = [YBookCode]::new($code)
    if ($run) {
        $yc.Run()
        return
    }
    return [PSCustomObject]@{
        "CheckDigit" = $yc.CheckDigit;
        "Full" = $yc.FullCode;
        "Format" = $yc.Format();
        "Url" = $yc.ToUrl();
    }
}

function yBookPage {
    param (
        [parameter(Mandatory)][string]$code
    )
    [YBookCode]::new($code).Run()
    Hide-ConsoleWindow
}

Set-PSReadLineKeyHandler -Key "ctrl+alt+y" -BriefDescription "ybookpage" -LongDescription "ybookpage" -ScriptBlock {
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert("yBookPage ")
}
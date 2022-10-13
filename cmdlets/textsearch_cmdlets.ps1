
<# ==============================

cmdlets for searching through text file

                encoding: utf8bom
============================== #>

function Select-StringHilight {
    <#
        .EXAMPLE
        ls -exclude *md | slh ほげ -encoding default
        ls | cat | slh ほげ
    #>
    [OutputType([System.Void])]
    param (
        [string]$pattern
        ,[switch]$case
        ,[int[]]$context = 0
        ,[ValidateSet("default", "oem")][string]$encoding = "default"
    )

    class GrepContext {
        static [void] Show ([string[]]$context, [int]$lineIndex, [bool]$post) {
            if (-not $context) {
                return
            }
            $l = ($post)? $lineIndex : $lineIndex - $context.Count - 1
            $context | ForEach-Object {
                $l += 1
                "{0:d4}:{1}" -f $l, $_ | Write-Host -ForegroundColor DarkGray
            }
        }
    }

    $grep = $input | Select-String -Encoding $encoding -Pattern $pattern -CaseSensitive:$case -AllMatches -Context $context
    foreach ($g in $grep) {
        [GrepContext]::Show($g.Context.PreContext, $g.LineNumber, $false)

        ($g.Filename -eq "InputStream")?
            "{0:d4}:" -f $g.LineNumber :
            "{0}:{1:d4}:" -f $g.Filename, $g.LineNumber | Write-Host -NoNewline -ForegroundColor DarkBlue

        $g.Line | hilight -pattern $pattern -case:$case -color "Yellow"

        [GrepContext]::Show($g.Context.PostContext, $g.LineNumber, $true)

    }
    $total = $grep.Matches.Count
    if ($total) {
        Write-Host ("========== {0} ==========" -f $total) -ForegroundColor Cyan
    }
}
Set-Alias slh Select-StringHilight

function Invoke-CountMatch {
    $pattern = $args | Join-String -Separator ")|(" -OutputPrefix "(" -OutputSuffix ")"
    $grep = $input | Where-Object {$_.trim()} | Select-String -Pattern $pattern -CaseSensitive -AllMatches
    return $($grep.Matches.Value | Sort-Object -CaseSensitive | Group-Object -NoElement)
}

function Get-MatchPattern {
    param (
        [string]$pattern,
        [switch]$case
    )
    $grep = @($input | Select-String -Pattern $pattern -AllMatches -CaseSensitive:$case)
    return $($grep.Matches.Value | Group-Object -NoElement | Sort-Object Count)
}

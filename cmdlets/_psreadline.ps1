
<# ==============================

PSReadLine Configuration

                encoding: utf8bom
============================== #>

using namespace Microsoft.PowerShell

# https://github.com/PowerShell/PSReadLine/blob/dc38b451be/PSReadLine/SamplePSReadLineProfile.ps1
# you can search keychar with `[Console]::ReadKey()`

Set-PSReadlineOption -HistoryNoDuplicates `
    -PredictionSource History `
    -BellStyle None `
    -ContinuationPrompt ($Global:PSStyle.Foreground.BrightBlack + "#>" + $Global:PSStyle.Reset) `
    -AddToHistoryHandler {
        param ($command)
        switch -regex ($command) {
            "SKIPHISTORY" {return $false}
            # "^[a-z]$" {return $false}
            # "^[a-z] " {return $false}
            # "exit" {return $false}
            "^dsk$" {return $false}
            " -execute" {return $false}
            " -force" {return $false}
        }
        return $true
    }

Set-PSReadLineOption -colors @{
    "Command" = $Global:PSStyle.Foreground.BrightYellow;
    "Comment" = $Global:PSStyle.Foreground.BrightBlack;
    "Number" = $Global:PSStyle.Foreground.BrightCyan;
    "String" = $Global:PSStyle.Foreground.BrightBlue;
    "Variable" = $Global:PSStyle.Foreground.BrightGreen;
    "InlinePrediction" = $Global:PSStyle.Foreground.Blue;
}

# history
Set-PSReadLineKeyHandler -Key "ctrl+r" -Function ReverseSearchHistory
Set-PSReadLineKeyHandler -Key "ctrl+R" -Function ForwardSearchHistory

# Shell cursor jump
Set-PSReadLineKeyHandler -Key "alt+j" -Function "ShellForwardWord"
Set-PSReadLineKeyHandler -Key "alt+k" -Function "ShellBackwardWord"
Set-PSReadLineKeyHandler -Key "alt+J" -Function "SelectShellForwardWord"
Set-PSReadLineKeyHandler -Key "alt+K" -Function "SelectShellBackwardWord"


# PS cursor jump
Set-PSReadLineOption -WordDelimiters ";:,.[]{}()/\|^&*-=+'`" !?@#`$%&_<>``「」（）『』『』［］、，。：；／　"
@{
    "ctrl+n" = "ForwardWord";
    "ctrl+RightArrow" = "ForwardWord";
    "ctrl+LeftArrow" = "BackwardWord";
    "ctrl+shift+RightArrow" = "SelectForwardWord";
    "ctrl+shift+LeftArrow" = "SelectBackwardWord";
}.GetEnumerator() | ForEach-Object {
    Set-PSReadLineKeyHandler -Key $_.Key -Function $_.Value
}
Set-PSReadLineKeyHandler -Key "ctrl+backspace" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    if ($rl.HasRange) {
        [PSConsoleReadLine]::BackwardDeleteChar()
    }
    [PSConsoleReadLine]::BackwardKillWord()
}
Set-PSReadLineKeyHandler -Key "ctrl+delete" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    if ($rl.HasRange) {
        [PSConsoleReadLine]::DeleteChar()
    }
    [PSConsoleReadLine]::KillWord()
}


class PSCursorLine {
    [string]$Text
    [string]$BeforeCursor
    [string]$AfterCursor
    [int]$Indent
    [int]$Index
    [int]$StartPos

    PSCursorLine([string]$line, [int]$pos) {
        $lines = $line -split "`n"
        $this.Index = ($line.Substring(0, $pos) -split "`n").Length - 1
        $this.Text = $lines[$this.Index]
        $this.Indent = $this.Text.Length - $this.Text.TrimStart().Length
        if ($this.Index -gt 0) {
            $this.StartPos = ($lines[0..($this.Index - 1)] -join "`n").Length + 1
        }
        else {
            $this.StartPos = 0
        }
        $this.BeforeCursor = $this.Text.Substring(0, $pos - $this.StartPos)
        $this.AfterCursor = $this.Text.Substring($pos - $this.StartPos)
    }

}

class ReadLiner2 {
    [string]$Commandline
    [int]$CursorPos
    [PSCursorLine]$CursorLine
    [int]$SelectionStart
    [int]$SelectionLength
    [bool]$HasRange

    ReadLiner2() {
        $line = $null
        $pos = $null
        [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$pos)
        $this.Commandline = $line
        $this.CursorPos = $pos
        $this.CursorLine = [PSCursorLine]::new($line, $pos)
        $start = $null
        $length = $null
        [PSConsoleReadLine]::GetSelectionState([ref]$start, [ref]$length)
        $this.SelectionStart = $start
        $this.SelectionLength = $length
        $this.HasRange = $this.SelectionLength -gt 0
    }

    [void] ToggleLineComment () {
        $pos = $this.CursorPos
        $top = $this.CursorLine.StartPos
        $indent = $this.CursorLine.Indent
        $curLine = $this.CursorLine.Text
        if ($curLine.TrimStart().StartsWith("#")) {
            $uncomment = $curLine -replace "(^ *)#", '$1'
            [PSConsoleReadLine]::Replace($top, $curLine.Length, $uncomment)
        }
        else {
            [PSConsoleReadLine]::SetCursorPosition($top + $indent)
            [PSConsoleReadLine]::Insert("#")
            [PSConsoleReadLine]::SetCursorPosition($pos + 1)
        }
    }

    [void] IndentLine () {
        $pos = $this.CursorPos
        $top = $this.CursorLine.StartPos
        [PSConsoleReadLine]::SetCursorPosition($top)
        [PSConsoleReadLine]::Insert("  ")
        [PSConsoleReadLine]::SetCursorPosition($pos+2)
    }

    [void] OutdentLine () {
        $indent = $this.CursorLine.Indent
        if ($indent -ge 2) {
            $top = $this.CursorLine.StartPos
            [PSConsoleReadLine]::Replace($top, $indent, " " * ($indent - 2))
            [PSConsoleReadLine]::SetCursorPosition($this.CursorPos - 2)
        }
    }

    [void] RemoveTrailingPipe () {
        $line = $this.Commandline
        $pos = $this.CursorPos
        if ($pos -gt 0) {
            $len = $line.Length - $line.TrimEnd(" |").Length
            [PSConsoleReadLine]::Delete($pos - $len, $len)
            $this.CursorPos = $this.CursorPos - $len
        }
    }

    [int] FindMatchingBracket () {
        $pairs = @{
            "{" = "}"; "[" = "]"; "(" = ")";
            "}" = "{"; "]" = "["; ")" = "(";
        }
        $line = $this.CommandLine
        $pos = $this.CursorPos
        $curChar = $line[$pos] -as [string]
        if ($curChar -notin $pairs.Keys) {
            return -1
        }
        $bracket = $curChar
        if ($bracket -in @("{", "[", "(")) {
            $max = $line.Length - $pos
            $step = 1
        }
        else {
            $max = $pos
            $step = -1
        }
        $target = $pairs[$bracket]
        $found = -1
        $skip = 0
        for ($i = $step; $i -lt $max; $i += $step) {
            $c = $line[$pos+$i] -as [string]
            if ($c -eq $bracket) {
                $skip += 1
                continue
            }
            if ($c -eq $target) {
                if ($skip -gt 0) {
                    $skip -= 1
                    continue
                }
                $found = $pos + $i
                break
            }
        }
        return $found
    }

}


class PsAst {
    $CommandAst = @()

    PsAst() {
        $ast = $null
        $tokens = $null
        $errors = $null
        $cursor = $null
        [PSConsoleReadLine]::GetBufferState([ref]$ast, [ref]$tokens, [ref]$errors, [ref]$cursor)
        $this.CommandAst = $ast.FindAll( {
            param ($params)
            return $params[0] -is [System.Management.Automation.Language.CommandAst]
        }, $true) | ForEach-Object {
            return [PSCustomObject]@{
                "Node" = $_;
                "CursorOn" = $($_.Extent.StartOffset -le $cursor) -and ($_.Extent.EndOffset -ge $cursor);
            }
        }
    }

}

Set-PSReadLineKeyHandler -Key "alt+q" -BriefDescription "test" -LongDescription "test" -ScriptBlock {
    $a = [PsAst]::new()
    $a.CommandAst | Write-Host
}

Set-PSReadLineKeyHandler -Key "ctrl+Q" -BriefDescription "exit" -LongDescription "exit" -ScriptBlock {
    [PSConsoleReadLine]::Insert("<#SKIPHISTORY#>exit")
}

Set-PSReadLineKeyHandler -Key "alt+[" -BriefDescription "insert-multiline-brace" -LongDescription "insert-multiline-brace" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.CursorPos
    $indent = $rl.CursorLine.Indent
    $filler = " " * $indent
    [PSConsoleReadLine]::Insert("{" + "`n  $filler`n$filler" + "}")
    [PSConsoleReadLine]::SetCursorPosition($pos + 2 + $indent + 2)
}

Set-PSReadLineKeyHandler -Key "ctrl+j" -Description "smart-InsertLineBelow" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $indent = $rl.CursorLine.Indent
    [PSConsoleReadLine]::InsertLineBelow()
    [PSConsoleReadLine]::Insert(" " * $indent)
}
Set-PSReadLineKeyHandler -Key "ctrl+J" -Description "smart-InsertLineAbove" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $indent = $rl.CursorLine.Indent
    [PSConsoleReadLine]::InsertLineAbove()
    [PSConsoleReadLine]::Insert(" " * $indent)
}

Set-PSReadLineKeyHandler -Key "Shift+Enter" -BriefDescription "addline-and-indent" -LongDescription "addline-and-indent" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos
    $indent = $rl.CursorLine.Indent
    $filler = " " * $indent
    if ($line[$pos] -eq "}") {
        if ($line[$pos - 1] -eq "{") {
            [PSConsoleReadLine]::Insert("`n" + $filler + "  `n" + $filler)
            [PSConsoleReadLine]::SetCursorPosition($pos + 1 + $filler.Length + 2)
            return
        }
        [PSConsoleReadLine]::Insert("`n" + $filler)
        return
    }
    if ($line[$pos - 1] -eq "{") {
            [PSConsoleReadLine]::Insert("`n" + "  " + $filler)
        return
    }
    [PSConsoleReadLine]::Insert("`n" + $filler)
}

# ctrl+shift+]
Set-PSReadLineKeyHandler -Key "ctrl+shift+Oem6" -BriefDescription "dupl-down" -LongDescription "duplicate-currentline-down" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $curLine = $rl.CursorLine.Text
    $curLineStart = $rl.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
}
# ctrl+shift+[
Set-PSReadLineKeyHandler -Key "ctrl+shift+Oem4" -BriefDescription "dupl-up" -LongDescription "duplicate-currentline-up" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $curLine = $rl.CursorLine.Text
    $curLineStart = $rl.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
    [PSConsoleReadLine]::SetCursorPosition($curLineStart)

}


# ctrl+]
Set-PSReadLineKeyHandler -Key "ctrl+Oem6" -BriefDescription "indent" -LongDescription "indent" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $rl.IndentLine()
}
# ctrl+[
Set-PSReadLineKeyHandler -Key "ctrl+Oem4" -BriefDescription "outdent" -LongDescription "outdent" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $rl.OutdentLine()
}

Set-PSReadLineKeyHandler -Key "ctrl+/" -BriefDescription "toggle-comment" -LongDescription "toggle-comment" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $rl.ToggleLineComment()
}

Set-PSReadLineKeyHandler -Key "alt+l" -BriefDescription "insert-pipe-and-adjust-pos" -LongDescription "insert-pipe-and-adjust-pos" -ScriptBlock {
    [PSConsoleReadLine]::Insert("|")
    $rl = [ReadLiner2]::new()
    if ($rl.CursorLine.AfterCursor.Trim().Length) {
        $pos = $rl.CursorPos
        if ($rl.CursorLine.AfterCursor.StartsWith("|")) {
            return
        }
        $len = $pos - ($rl.CursorLine.BeforeCursor.TrimEnd(" |").Length)
        [PSConsoleReadLine]::Replace($pos - $len, $len, "||")
        [PSConsoleReadLine]::SetCursorPosition($pos - $len + 1)
    }
}

Set-PSReadLineKeyHandler -Key "ctrl+|" -BriefDescription "find-matching-bracket" -LongDescription "find-matching-bracket" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.FindMatchingBracket()
    if ($pos -ge 0) {
        [PSConsoleReadLine]::SetCursorPosition($pos)
    }
}
Set-PSReadLineKeyHandler -Key "alt+P" -BriefDescription "remove-matchingBraces" -LongDescription "remove-matchingBraces" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.FindMatchingBracket()
    if ($pos -ge 0) {
        $start = [math]::Min($pos, $rl.CursorPos)
        $end = [math]::Max($pos, $rl.CursorPos)
        $len = $end - $start + 1
        $repl = $rl.CommandLine.Substring($start+1, $len-2)
        [PSConsoleReadLine]::Replace($start, $len, $repl)
    }
}

##############################
# smart home
##############################

Set-PSReadLineKeyHandler -Key "home" -BriefDescription "smart-home" -LongDescription "smart-home" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.CursorPos
    $indent = $rl.CursorLine.Indent
    $top = $rl.CursorLine.StartPos
    if ($pos -gt $top + $indent) {
        [PSConsoleReadLine]::SetCursorPosition($top + $indent)
    }
    else {
        [PSConsoleReadLine]::BeginningOfLine()
    }
}

##############################
# smart paste
##############################

Set-PSReadLineKeyHandler -Key "ctrl+k,v" -BriefDescription "smart-paste" -LongDescription "smart-paste" -ScriptBlock {
    $cb = @(Get-Clipboard)
    $s = ($cb.Count -gt 1)? ($cb | ForEach-Object {($_ -as [string]).TrimEnd()} | Join-String -Separator "`n" -OutputPrefix "@'`n" -OutputSuffix "`n'@") : $cb
    if ([ReadLiner2]::new().HasRange) {
        [PSConsoleReadLine]::DeleteChar()
    }
    [PSConsoleReadLine]::Insert($s)
}

##############################
# yank last argument as variable
##############################

Set-PSReadLineKeyHandler -Key "alt+a" -BriefDescription "yankLastArgAsVariable" -LongDescription "yankLastArgAsVariable" -ScriptBlock {
    [PSConsoleReadLine]::Insert("$")
    [PSConsoleReadLine]::YankLastArg()
    $line = [ReadLiner2]::new().CommandLine
    if ($line -match '\$\$') {
        $newLine = $line -replace '\$\$', "$"
        [PSConsoleReadLine]::Replace(0, $line.Length, $newLine)
    }
}

##############################
# smart brackets
##############################

Set-PSReadLineKeyHandler -Key "(","{","[" -BriefDescription "InsertPairedBraces" -LongDescription "Insert matching braces or wrap selection by matching braces" -ScriptBlock {
    param($key, $arg)
    $openChar = $key.KeyChar
    $closeChar = switch ($openChar) {
        <#case#> "(" { [char]")"; break }
        <#case#> "{" { [char]"}"; break }
        <#case#> "[" { [char]"]"; break }
    }

    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos

    if ($rl.HasRange) {
        [PSConsoleReadLine]::Replace($rl.SelectionStart, $rl.selectionLength, $openChar + $line.SubString($rl.selectionStart, $rl.selectionLength) + $closeChar)
        [PSConsoleReadLine]::SetCursorPosition($rl.selectionStart + $rl.selectionLength + 2)
        return
    }

    $nOpen = [regex]::Matches($line, [regex]::Escape($openChar)).Count
    $nClose = [regex]::Matches($line, [regex]::Escape($closeChar)).Count
    if ($nOpen -ne $nClose) {
        [PSConsoleReadLine]::Insert($openChar)
        return
    }
    [PSConsoleReadLine]::Insert($openChar + $closeChar)
    [PSConsoleReadLine]::SetCursorPosition($pos + 1)
}

Set-PSReadLineKeyHandler -Key ")","]","}" -BriefDescription "SmartCloseBraces" -LongDescription "Insert closing brace or skip" -ScriptBlock {
    param($key, $arg)

    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos

    if ($line[$pos] -eq $key.KeyChar) {
        [PSConsoleReadLine]::SetCursorPosition($pos + 1)
    }
    else {
        [PSConsoleReadLine]::Insert($key.KeyChar)
    }
}

Set-PSReadLineKeyHandler -Key "alt+w","alt+(" -BriefDescription "WrapLineByParenthesis" -LongDescription "Wrap the entire line or selection and move the cursor after the closing punctuation" -ScriptBlock {
    $prefix, $suffix = @('(', ')')
    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    if ($rl.HasRange) {
        [PSConsoleReadLine]::Replace($rl.selectionStart, $rl.selectionLength, $prefix + $line.SubString($rl.selectionStart, $rl.selectionLength) + $suffix)
        [PSConsoleReadLine]::SetCursorPosition($rl.selectionStart + $rl.selectionLength + 2)
        return
    }
    $curLine = $rl.CursorLine.Text
    $curLineStart = $rl.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $prefix + $curLine + $suffix)
}

Set-PSReadLineKeyHandler -Key "ctrl+k,t" -BriefDescription "cast-as-type" -LongDescription "cast-as-type" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    if (-not $rl.HasRange) {
        return
    }
    $repl = "({0} -as [])" -f $line.SubString($rl.selectionStart, $rl.selectionLength)
    [PSConsoleReadLine]::Replace($rl.selectionStart, $rl.selectionLength, $repl)
    [PSConsoleReadLine]::SetCursorPosition($rl.selectionStart + $rl.selectionLength + 7)
}

##############################
# smart method completion
##############################

Remove-PSReadlineKeyHandler "tab"
Set-PSReadLineKeyHandler -Key "tab" -BriefDescription "smartNextCompletion" -LongDescription "insert closing parenthesis in forward completion of method" -ScriptBlock {

    [PSConsoleReadLine]::TabCompleteNext()

    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos
    if ($line[($pos - 1)] -eq "(") {
        if ($line[$pos] -ne ")") {
            [PSConsoleReadLine]::Insert(")")
            [PSConsoleReadLine]::BackwardChar()
        }
    }
}

Remove-PSReadlineKeyHandler "shift+tab"
Set-PSReadLineKeyHandler -Key "shift+tab" -BriefDescription "smartPreviousCompletion" -LongDescription "insert closing parenthesis in backward completion of method" -ScriptBlock {

    [PSConsoleReadLine]::TabCompletePrevious()
    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos

    if ($line[($pos - 1)] -eq "(") {
        if ($line[$pos] -ne ")") {
            [PSConsoleReadLine]::Insert(")")
            [PSConsoleReadLine]::BackwardChar()
        }
    }
}

##############################
# smart quotation
##############################

Set-PSReadLineKeyHandler -Key "`"","'" -BriefDescription "smartQuotation" -LongDescription "Put quotation marks and move the cursor between them or put marks around the selection" -ScriptBlock {
    param($key, $arg)
    $mark = $key.KeyChar

    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos

    if ($rl.HasRange) {
        [PSConsoleReadLine]::Replace($rl.selectionStart, $rl.selectionLength, $mark + $line.SubString($rl.selectionStart, $rl.selectionLength) + $mark)
        [PSConsoleReadLine]::SetCursorPosition($rl.selectionStart + $rl.selectionLength + 2)
        return
    }

    if ($line[$pos] -eq $mark) {
        [PSConsoleReadLine]::SetCursorPosition($pos + 1)
        return
    }

    $nMark = [regex]::Matches($line, $mark).Count
    if ($nMark % 2 -eq 1) {
        [PSConsoleReadLine]::Insert($mark)
    }
    else {
        [PSConsoleReadLine]::Insert($mark + $mark)
        [PSConsoleReadLine]::SetCursorPosition($pos+1)
    }
}

##############################
# reload profile
##############################

Set-PSReadLineKeyHandler -Key "alt+r" -BriefDescription "reloadPROFILE" -LongDescription "reloadPROFILE" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert('<#SKIPHISTORY#> . $PROFILE')
    [PSConsoleReadLine]::AcceptLine()
}


##############################
# snippets
##############################

Set-PSReadLineKeyHandler -Key "alt+R,p" -BriefDescription "insert-regex-paren" -LongDescription "insert-regex-paren" -ScriptBlock {
    [PSConsoleReadLine]::Insert('[\(（].+?[\)）]')
}
Set-PSReadLineKeyHandler -Key "alt+R,b" -BriefDescription "insert-regex-bracket" -LongDescription "insert-regex-bracket" -ScriptBlock {
    [PSConsoleReadLine]::Insert('[\[［].+?[\]］]')
}

Set-PSReadLineKeyHandler -Key "ctrl+k,p" -BriefDescription "psobject-name" -LongDescription "psobject-name" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.CursorPos
    if ($rl.Commandline[$pos - 1] -ne ".") {
        [PSConsoleReadLine]::Insert('.')
    }
    [PSConsoleReadLine]::Insert('psobject.properties.name')
}

Set-PSReadLineKeyHandler -Key "ctrl+k,i" -BriefDescription "insert-if-else-block" -LongDescription "insert-if-else-block" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $pos = $rl.CursorPos
    $indent = $rl.CursorLine.Indent
    $filler = " " * $indent
    $lines = @('if ($_ ) {', ($filler + '  $_'), ($filler + "}"), ($filler + "else {"), ($filler + '  $_'), ($filler + "}"))
    [PSConsoleReadLine]::Insert($lines -join "`n")
    [PSConsoleReadLine]::SetCursorPosition($pos + 7)
}

Set-PSReadLineKeyHandler -Key "alt+B" -BriefDescription "insert-scriptblock" -LongDescription "insert-scriptblock" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert('{$_}')
    [PSConsoleReadLine]::BackwardChar()
}
Set-PSReadLineKeyHandler -Key "ctrl+k,f","ctrl+k,w" -BriefDescription "insert-alias" -LongDescription "insert-alias" -ScriptBlock {
    param ($key, $arg)
    $a = switch ($key.KeyChar) {
        "f" { "%"; break }
        "w" { "?"; break }
    }
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|{0} " -f $a)
}

Set-PSReadLineKeyHandler -Key "ctrl+\" -BriefDescription "file-completion" -LongDescription "file-completion" -ScriptBlock {
    [PSConsoleReadLine]::Insert(".\")
    [PSConsoleReadLine]::MenuComplete()
}

Set-PSReadLineKeyHandler -Key "ctrl+V" -BriefDescription "setClipString" -LongDescription "setClipString" -ScriptBlock {
    $command = '<#SKIPHISTORY#> (gcb -Raw).Replace("`r","").Trim() -split "`n"|sv CLIPPING'
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert($command)
    [PSConsoleReadLine]::AddToHistory('$CLIPPING ')
    [PSConsoleReadLine]::AcceptLine()
}


Set-PSReadLineKeyHandler -Key "alt+n" -BriefDescription "filterEmpty" -LongDescription "filterEmpty" -ScriptBlock {
    $rl = [ReadLiner2]::new()
    $line = $rl.CommandLine
    $pos = $rl.CursorPos
    $s = ($line[$pos - 1] -eq "|")? '?{$_}|' : '|?{$_}'
    [PSConsoleReadLine]::Insert($s)
}

Set-PSReadLineKeyHandler -Key "alt+m" -BriefDescription "measure" -LongDescription "measure" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|measure ")
}

Set-PSReadLineKeyHandler -Key "alt+c" -BriefDescription "copyToClipboard" -LongDescription "copyToClipboard" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|c")
}

Set-PSReadLineKeyHandler -Key "alt+v" -BriefDescription "asVariable" -LongDescription "asVariable" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|v ")
}

Set-PSReadLineKeyHandler -Key "alt+t","alt+V" -BriefDescription "teeVariable" -LongDescription "teeVariable" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|tee -Variable ")
}

Set-PSReadLineKeyHandler -Key "alt+b" -BriefDescription "bat-plain" -LongDescription "bat-plain" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert("|oss|bat -p")
}

Set-PSReadLineKeyHandler -Key "alt+i" -BriefDescription "insert-invoke" -LongDescription "insert-invoke" -ScriptBlock {
    [PSConsoleReadLine]::Insert("Invoke*")
}

Set-PSReadLineKeyHandler -Key "ctrl+k,s" -BriefDescription "insert-Select-Object" -LongDescription "insert-Select-Object" -ScriptBlock {
    [ReadLiner2]::new().RemoveTrailingPipe()
    [PSConsoleReadLine]::Insert('|select -')
    [PSConsoleReadLine]::MenuComplete()
}

Set-PSReadLineKeyHandler -Key "alt+R,p","alt+R,9","alt+R,8" -BriefDescription "regexp-insideParen" -LongDescription "regexp-insideParen" -ScriptBlock {
    param($key, $arg)
    $reg = switch ($key.KeyChar) {
        <#case#> "8" { "（.+?）"; break }
        <#case#> "9" { "\(.+?\)"; break }
        <#case#> "p" { "[\(（].+?[\)）]"; break }
    }
    [PSConsoleReadLine]::Insert($reg)
}

# Set-PSReadLineKeyHandler -Key "alt+R,f" -BriefDescription "replaceWithFunction" -LongDescription "replaceWithFunction" -ScriptBlock {
#     [PSConsoleReadLine]::Insert('[regex]::Replace("", "", {$args[0].Value})')
# }

# Set-PSReadLineKeyHandler -Key "alt+R,g" -BriefDescription "replaceWithFunction-grouping" -LongDescription "replaceWithFunction-grouping" -ScriptBlock {
#     [PSConsoleReadLine]::Insert('[regex]::Replace("", "", {$args[0].groups[1].Value})')
# }

Set-PSReadLineKeyHandler -Key "alt+0","alt+-" -BriefDescription "insertAsterisk(star)" -LongDescription "insertAsterisk(star)" -ScriptBlock {
    [PSConsoleReadLine]::Insert("*")
}

##############################
# open folder
##############################

# desktop
Set-PSReadLineKeyHandler -Key "ctrl+d" -BriefDescription "desktop" -LongDescription "Invoke desktop" -ScriptBlock {
    $tablacus = "{0}\Dropbox\portable_apps\tablacus\TE64.exe" -f $env:USERPROFILE
    if (Test-Path $tablacus) {
        Start-Process $tablacus -ArgumentList @("{0}\desktop" -f $env:USERPROFILE)
    }
    else {
        Start-Process ("{0}\desktop" -f $env:USERPROFILE)
    }
    Hide-ConsoleWindow
}

# scan folder
Set-PSReadLineKeyHandler -Key "ctrl+S" -BriefDescription "openScanFolder" -LongDescription "openScanFolder" -ScriptBlock {
    $scanDir = "X:\scan"
    if(Test-Path $scanDir) {
        $tablacus = "{0}\Dropbox\portable_apps\tablacus\TE64.exe" -f $env:USERPROFILE
        if (Test-Path $tablacus) {
            Start-Process $tablacus -ArgumentList $scanDir
        }
        else {
            Start-Process $scanDir
        }
        Hide-ConsoleWindow
    }
}

################################
# AST
################################

# https://github.com/pecigonzalo/Oh-My-Posh/blob/master/plugins/psreadline/psreadline.ps1
Set-PSReadlineKeyHandler -Key "alt+h" -BriefDescription "CommandHelp" -LongDescription "Open the help window for the current command" -ScriptBlock {
    param($key, $arg)

    $ast = $null
    $tokens = $null
    $errors = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$ast, [ref]$tokens, [ref]$errors, [ref]$cursor)

    $commandAst = $ast.FindAll( {
        $node = $args[0]
        $node -is [System.Management.Automation.Language.CommandAst] -and
        $node.Extent.StartOffset -le $cursor -and
        $node.Extent.EndOffset -ge $cursor
    }, $true) | Select-Object -Last 1

    if ($commandAst -ne $null) {
        $commandName = $commandAst.GetCommandName()
        if ($commandName -ne $null) {
            $command = $ExecutionContext.InvokeCommand.GetCommand($commandName, 'All')
            if ($command -is [System.Management.Automation.AliasInfo]) {
            $commandName = $command.ResolvedCommandName
            }

            if ($commandName -ne $null) {
                Get-Help $commandName -ShowWindow
            }
        }
    }
}

Set-PSReadlineKeyHandler -Key "alt+L" -BriefDescription "toPreviousPipe" -LongDescription "toPreviousPipe" -ScriptBlock {
    param($key, $arg)

    $ast = $null
    $tokens = $null
    $errors = $null
    $cursor = $null
    [PSConsoleReadLine]::GetBufferState([ref]$ast, [ref]$tokens, [ref]$errors, [ref]$cursor)

    $lastPipe = $tokens | Where-Object {$_.Kind -eq "Pipe"} | Where-Object {$_.Extent.EndOffset -lt $cursor} | Select-Object -Last 1
    if ($lastPipe) {
        [PSConsoleReadLine]::SetCursorPosition($lastPipe.Extent.EndOffset - 1)
    }

}

##############################
# others
##############################

# format string
Set-PSReadLineKeyHandler -Key "ctrl+k,0", "ctrl+k,1", "ctrl+k,2", "ctrl+k,3", "ctrl+k,4", "ctrl+k,5", "ctrl+k,6", "ctrl+k,7", "ctrl+k,8", "ctrl+k,9" -ScriptBlock {
    param($key, $arg)
    $str = '{' + $key.KeyChar + '}'
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert($str)
}

# open from clipboard path
function ccat ([string]$encoding = "utf8") {
    $clip = (Get-Clipboard | Select-Object -First 1) -replace '"'
    if (Test-Path $clip -PathType Leaf) {
        Get-Content $clip -Encoding $encoding
    }
    else {
        "invalid-path!" | Write-Host -ForegroundColor Magenta
    }
}
Set-PSReadLineKeyHandler -Key "ctrl+p" -BriefDescription "setClipString" -LongDescription "setClipString" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert("ccat ")
}

Set-PSReadLineKeyHandler -Key "alt+d" -BriefDescription "openDraft" -LongDescription "openDraft" -ScriptBlock {
    $draft = "C:\Users\{0}\Dropbox\draft.txt" -f $env:USERNAME
    if (Test-Path $draft -PathType Leaf) {
        Start-Process $draft
        Hide-ConsoleWindow
    }
}

# pip
Set-PSReadLineKeyHandler -Key "alt+p,i", "alt+p,u", "alt+p,o" -BriefDescription "python-pip" -LongDescription "python-pip" -ScriptBlock {
    param($key, $arg)

    $opt = switch ($key.KeyChar) {
        <#case#> "i" {'install '; break}
        <#case#> "u" {'install --upgrade '; break}
        <#case#> "o" {'list --outdated'; break}
    }
    $command = 'python -m pip {0}' -f $opt
    [Microsoft.PowerShell.PSConsoleReadLine]::RevertLine()
    [Microsoft.PowerShell.PSConsoleReadLine]::Insert($command)
}
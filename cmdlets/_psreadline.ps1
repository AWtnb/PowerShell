﻿
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

# exit
Set-PSReadLineKeyHandler -Key "ctrl+Q" -BriefDescription "exit" -LongDescription "exit" -ScriptBlock {
    [PSConsoleReadLine]::Insert("<#SKIPHISTORY#>exit")
}

# reload
Set-PSReadLineKeyHandler -Key "alt+r" -BriefDescription "reloadPROFILE" -LongDescription "reloadPROFILE" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert('<#SKIPHISTORY#> . $PROFILE')
    [PSConsoleReadLine]::AcceptLine()
}

# completion
Set-PSReadLineKeyHandler -Key "alt+i" -BriefDescription "insert-invoke" -LongDescription "insert-invoke" -ScriptBlock {
    [PSConsoleReadLine]::Insert("Invoke*")
}
Set-PSReadLineKeyHandler -Key "alt+0","alt+-" -BriefDescription "insertAsterisk(star)" -LongDescription "insertAsterisk(star)" -ScriptBlock {
    [PSConsoleReadLine]::Insert("*")
}

# load clipboard
Set-PSReadLineKeyHandler -Key "ctrl+V" -BriefDescription "setClipString" -LongDescription "setClipString" -ScriptBlock {
    $command = '<#SKIPHISTORY#> (gcb -Raw).Replace("`r","").Trim() -split "`n"|sv CLIPPING'
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert($command)
    [PSConsoleReadLine]::AddToHistory('$CLIPPING ')
    [PSConsoleReadLine]::AcceptLine()
}

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

# open draft
Set-PSReadLineKeyHandler -Key "alt+d" -BriefDescription "openDraft" -LongDescription "openDraft" -ScriptBlock {
    $draft = "C:\Users\{0}\Dropbox\draft.txt" -f $env:USERNAME
    if (Test-Path $draft -PathType Leaf) {
        Start-Process $draft
        Hide-ConsoleWindow
    }
}

# cursor jump
Set-PSReadLineOption -WordDelimiters ";:,.[]{}()/\|^&*-=+'`" !?@#`$%&_<>``「」（）『』『』［］、，。：；／　"
@{
    "ctrl+RightArrow" = "ForwardWord";
    "ctrl+LeftArrow" = "BackwardWord";
    "ctrl+shift+RightArrow" = "SelectForwardWord";
    "ctrl+shift+LeftArrow" = "SelectBackwardWord";
}.GetEnumerator() | ForEach-Object {
    Set-PSReadLineKeyHandler -Key $_.Key -Function $_.Value
}

# https://github.com/pecigonzalo/Oh-My-Posh/blob/master/plugins/psreadline/psreadline.ps1
class ASTer {
    $ast
    $tokens
    $errors
    $cursor
    ASTer() {
        $a = $t = $e = $c = $null
        [PSConsoleReadLine]::GetBufferState([ref]$a, [ref]$t, [ref]$e, [ref]$c)
        $this.ast = $a
        $this.tokens = $t
        $this.errors = $e
        $this.cursor = $c
    }

    [System.Management.Automation.Language.Ast[]] Listup([string]$name) {
        return $this.ast.FindAll({
            return $args[0].GetType().Name.EndsWith($name)
        }, $true)
    }

    [System.Management.Automation.Language.Ast] GetActiveAst([string]$name) {
        return $this.Listup($name) | Where-Object {
            return ($_.Extent.StartOffset -le $this.cursor) -and ($this.cursor -le $_.Extent.EndOffset)
        } | Select-Object -Last 1
    }

    [System.Management.Automation.Language.Ast] GetPreviousAst([string]$name) {
        return $this.Listup($name) | Where-Object { $_.Extent.EndOffset -lt $this.cursor } | Select-Object -Last 1
    }

    [System.Management.Automation.Language.Ast] GetNextAst([string]$name) {
        return $this.Listup($name) | Where-Object { $_.Extent.StartOffset -gt $this.cursor } | Select-Object -First 1
    }


    [int] GetActiveTokenIndex() {
        $idx = -1
        foreach ($token in $this.tokens) {
            $idx += 1
            if (($token.Extent.StartOffset -le $this.cursor) -and ($this.cursor -le $token.Extent.EndOffset)) {
                break;
            }
        }
        return $idx
    }

    [System.Management.Automation.Language.Token] GetActiveToken() {
        $i = $this.GetActiveTokenIndex()
        return $this.tokens[$i]
        }

    [System.Management.Automation.Language.Token] GetPreviousToken() {
        $i = $this.GetActiveTokenIndex()
        $pos = $i - 1
        return $this.tokens[$pos]
    }
    [System.Management.Automation.Language.Token] GetNextToken() {
        $i = $this.GetActiveTokenIndex()
        $pos = $i + 1
        return $this.tokens[$pos]
    }

    [bool] IsStartOfToken() {
        return $this.cursor -eq $this.GetActiveToken().Extent.StartOffset
    }

    [bool] IsEndOfToken() {
        return $this.cursor -eq $this.GetActiveToken().Extent.EndOffset
    }

    [bool] IsAfterPipe() {
        if ($this.IsEndOfToken()) {
            return $this.GetActiveToken().Kind -eq "Pipe"
        }
        return $this.GetPreviousToken().Kind -eq "Pipe"
    }

}


Set-PSReadlineKeyHandler -Key "ctrl+alt+l" -BriefDescription "toPreviousPipe" -LongDescription "toPreviousPipe" -ScriptBlock {
    $a = [ASTer]::new()
    $lastPipe = $a.tokens | Where-Object {$_.Kind -eq "Pipe"} | Where-Object {$_.Extent.EndOffset -lt $a.cursor} | Select-Object -Last 1
    if ($lastPipe) {
        [PSConsoleReadLine]::SetCursorPosition($lastPipe.Extent.EndOffset - 1)
    }
}

Set-PSReadLineKeyHandler -Key "alt+l" -BriefDescription "insert-pipe" -LongDescription "insert-pipe" -ScriptBlock {
    $a = [ASTer]::new()
    [PSConsoleReadLine]::Insert("|")
    if ($a.IsAfterPipe()) {
        [PSConsoleReadLine]::BackwardChar()
    }
}

Set-PSReadLineKeyHandler -Key "ctrl+k,l" -BriefDescription "insert-pipe-to-head" -LongDescription "insert-pipe-to-head" -ScriptBlock {
    $a = [ASTer]::new()
    $activeCmd = $a.GetActiveAst("CommandAst")
    if ($activeCmd) {
        [PSConsoleReadLine]::SetCursorPosition($activeCmd.Extent.StartOffset)
        [PSConsoleReadLine]::Insert("|")
        [PSConsoleReadLine]::BackwardChar()
        return
    }
    $lastCmd = $a.GetPreviousAst("CommandAst")
    if ($lastCmd) {
        [PSConsoleReadLine]::SetCursorPosition($lastCmd.Extent.StartOffset)
    }
    [PSConsoleReadLine]::Insert("|")
    [PSConsoleReadLine]::BackwardChar()
}

Set-PSReadLineKeyHandler -Key "ctrl+k,alt+l" -BriefDescription "insert-pipe-to-tail" -LongDescription "insert-pipe-to-tail" -ScriptBlock {
    $a = [ASTer]::new()
    $activeCmd = $a.GetActiveAst("CommandAst")
    if ($activeCmd) {
        [PSConsoleReadLine]::SetCursorPosition($activeCmd.Extent.EndOffset)
        [PSConsoleReadLine]::Insert("|")
        return
    }
    $nextCmd = $a.GetNextAst("CommandAst")
    if ($nextCmd) {
        [PSConsoleReadLine]::SetCursorPosition($nextCmd.Extent.EndOffset)
    }
    [PSConsoleReadLine]::Insert("|")
}


Set-PSReadLineKeyHandler -Key "ctrl+n" -BriefDescription "smart-forwardWord" -LongDescription "smart-forwardWord" -ScriptBlock {
    [PSConsoleReadLine]::ForwardWord()
    $a = [ASTer]::new()
    $aToken = $a.GetActiveToken()
    if (
        ($aToken.Kind -eq "StringExpandable" -and -not $aToken.Text.EndsWith('"')) `
        -or `
        ($aToken.Kind -eq "StringLiteral" -and -not $aToken.Text.EndsWith("'")) `
    ) {
        [PSConsoleReadLine]::Insert($aToken.Text.Substring(0,1))
    }
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

class PSBufferState {
    [string]$Commandline
    [int]$CursorPos
    [PSCursorLine]$CursorLine
    [int]$SelectionStart
    [int]$SelectionLength

    PSBufferState() {
        $line = $pos = $null
        [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$pos)
        $this.Commandline = $line
        $this.CursorPos = $pos
        $this.CursorLine = [PSCursorLine]::new($line, $pos)
        $start = $length = $null
        [PSConsoleReadLine]::GetSelectionState([ref]$start, [ref]$length)
        $this.SelectionStart = $start
        $this.SelectionLength = $length
    }

    [void] ToggleLineComment() {
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

    [void] IndentLine() {
        $pos = $this.CursorPos
        $top = $this.CursorLine.StartPos
        [PSConsoleReadLine]::SetCursorPosition($top)
        [PSConsoleReadLine]::Insert("  ")
        [PSConsoleReadLine]::SetCursorPosition($pos+2)
    }

    [void] OutdentLine() {
        $indent = $this.CursorLine.Indent
        if ($indent -ge 2) {
            $top = $this.CursorLine.StartPos
            [PSConsoleReadLine]::Replace($top, $indent, " " * ($indent - 2))
            [PSConsoleReadLine]::SetCursorPosition($this.CursorPos - 2)
        }
    }

    static [bool] IsSelecting() {
        $start = $length = $null
        [PSConsoleReadLine]::GetSelectionState([ref]$start, [ref]$length)
        return $length -gt 0
    }

    static [int] FindMatchingPairPos() {
        $a = [ASTer]::new()
        $activeToken = $a.GetActiveToken()
        if (-not $activeToken) {
            return -1
        }
        if ($activeToken.Kind.ToString().Substring(1) -notin @("Bracket", "Curly", "Paren")) {
            return -1
        }
        $cur = $activeToken.Kind.ToString()
        $name = $cur.Substring(1)
        $skip = 0
        $goal = -1
        $focus = $a.tokens | Where-Object {$_.Kind.ToString().EndsWith($name)}
        if ($cur.StartsWith("L")) {
            $targets = $focus | Where-Object {$_.Extent.StartOffset -ge $activeToken.Extent.EndOffset}
        }
        else {
            $targets = $focus | Where-Object {$_.Extent.EndOffset -le $activeToken.Extent.StartOffset} | Sort-Object -Descending {$_.Extent.StartOffset}
        }

        foreach ($token in $targets) {
            if ($token.Kind.ToString() -eq $cur) {
                $skip += 1
                continue
            }
            if ($skip -gt 0) {
                $skip += -1
                continue
            }
            $goal = $token.Extent.StartOffset
            break
        }
        return $goal
    }

}

Set-PSReadLineKeyHandler -Key "ctrl+backspace" -ScriptBlock {
    if ([PSBufferState]::IsSelecting()) {
        [PSConsoleReadLine]::BackwardDeleteChar()
    }
    [PSConsoleReadLine]::BackwardKillWord()
}
Set-PSReadLineKeyHandler -Key "ctrl+delete" -ScriptBlock {
    if ([PSBufferState]::IsSelecting()) {
        [PSConsoleReadLine]::DeleteChar()
    }
    [PSConsoleReadLine]::KillWord()
}


Set-PSReadLineKeyHandler -Key "alt+[" -BriefDescription "insert-multiline-brace" -LongDescription "insert-multiline-brace" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $pos = $bs.CursorPos
    $indent = $bs.CursorLine.Indent
    $filler = " " * $indent
    [PSConsoleReadLine]::Insert("{" + "`n  $filler`n$filler" + "}")
    [PSConsoleReadLine]::SetCursorPosition($pos + 2 + $indent + 2)
}

Set-PSReadLineKeyHandler -Key "ctrl+j" -Description "smart-InsertLineBelow" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $indent = $bs.CursorLine.Indent
    [PSConsoleReadLine]::InsertLineBelow()
    [PSConsoleReadLine]::Insert(" " * $indent)
}
Set-PSReadLineKeyHandler -Key "ctrl+J" -Description "smart-InsertLineAbove" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $indent = $bs.CursorLine.Indent
    [PSConsoleReadLine]::InsertLineAbove()
    [PSConsoleReadLine]::Insert(" " * $indent)
}

Set-PSReadLineKeyHandler -Key "Shift+Enter" -BriefDescription "addline-and-indent" -LongDescription "addline-and-indent" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    $pos = $bs.CursorPos
    $indent = $bs.CursorLine.Indent
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
    $bs = [PSBufferState]::new()
    $curLine = $bs.CursorLine.Text
    $curLineStart = $bs.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
}
# ctrl+shift+[
Set-PSReadLineKeyHandler -Key "ctrl+shift+Oem4" -BriefDescription "dupl-up" -LongDescription "duplicate-currentline-up" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $curLine = $bs.CursorLine.Text
    $curLineStart = $bs.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
    [PSConsoleReadLine]::SetCursorPosition($curLineStart)

}


# ctrl+]
Set-PSReadLineKeyHandler -Key "ctrl+Oem6" -BriefDescription "indent" -LongDescription "indent" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.IndentLine()
}
# ctrl+[
Set-PSReadLineKeyHandler -Key "ctrl+Oem4" -BriefDescription "outdent" -LongDescription "outdent" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.OutdentLine()
}

Set-PSReadLineKeyHandler -Key "ctrl+/" -BriefDescription "toggle-comment" -LongDescription "toggle-comment" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.ToggleLineComment()
}

Set-PSReadLineKeyHandler -Key "ctrl+|" -BriefDescription "find-matching-bracket" -LongDescription "find-matching-bracket" -ScriptBlock {
    $pos = [PSBufferState]::FindMatchingPairPos()
    if ($pos -ge 0) {
        [PSConsoleReadLine]::SetCursorPosition($pos)
    }
}
Set-PSReadLineKeyHandler -Key "alt+P" -BriefDescription "remove-matchingBraces" -LongDescription "remove-matchingBraces" -ScriptBlock {
    $pos = [PSBufferState]::FindMatchingPairPos()
    if ($pos -ge 0) {
        $start = [math]::Min($pos, $bs.CursorPos)
        $end = [math]::Max($pos, $bs.CursorPos)
        $len = $end - $start + 1
        $repl = $bs.CommandLine.Substring($start+1, $len-2)
        [PSConsoleReadLine]::Replace($start, $len, $repl)
    }
}

Set-PSReadLineKeyHandler -Key "home" -BriefDescription "smart-home" -LongDescription "smart-home" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $pos = $bs.CursorPos
    $indent = $bs.CursorLine.Indent
    $top = $bs.CursorLine.StartPos
    if ($pos -gt $top + $indent) {
        [PSConsoleReadLine]::SetCursorPosition($top + $indent)
    }
    else {
        [PSConsoleReadLine]::BeginningOfLine()
    }
}


Set-PSReadLineKeyHandler -Key "ctrl+k,v" -BriefDescription "smart-paste" -LongDescription "smart-paste" -ScriptBlock {
    $cb = @(Get-Clipboard)
    $s = ($cb.Count -gt 1)? ($cb | ForEach-Object {($_ -as [string]).TrimEnd()} | Join-String -Separator "`n" -OutputPrefix "@'`n" -OutputSuffix "`n'@") : $cb
    if ([PSBufferState]::IsSelecting()) {
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
    $line = [PSBufferState]::new().CommandLine
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

    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    $pos = $bs.CursorPos

    if ($bs.SelectionLength -gt 0) {
        [PSConsoleReadLine]::Replace($bs.SelectionStart, $bs.selectionLength, $openChar + $line.SubString($bs.selectionStart, $bs.selectionLength) + $closeChar)
        [PSConsoleReadLine]::SetCursorPosition($bs.selectionStart + $bs.selectionLength + 2)
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

    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    $pos = $bs.CursorPos

    if ($line[$pos] -eq $key.KeyChar) {
        [PSConsoleReadLine]::SetCursorPosition($pos + 1)
    }
    else {
        [PSConsoleReadLine]::Insert($key.KeyChar)
    }
}

Set-PSReadLineKeyHandler -Key "alt+w","alt+(" -BriefDescription "WrapLineByParenthesis" -LongDescription "Wrap the entire line or selection and move the cursor after the closing punctuation" -ScriptBlock {
    $prefix, $suffix = @('(', ')')
    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    if ($bs.SelectionLength -gt 0) {
        [PSConsoleReadLine]::Replace($bs.selectionStart, $bs.selectionLength, $prefix + $line.SubString($bs.selectionStart, $bs.selectionLength) + $suffix)
        [PSConsoleReadLine]::SetCursorPosition($bs.selectionStart + $bs.selectionLength + 2)
        return
    }
    $curLine = $bs.CursorLine.Text
    $curLineStart = $bs.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $prefix + $curLine + $suffix)
}

Set-PSReadLineKeyHandler -Key "ctrl+k,t" -BriefDescription "cast-as-type" -LongDescription "cast-as-type" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    if (-not $bs.SelectionLength -gt 0) {
        return
    }
    $repl = "({0} -as [])" -f $line.SubString($bs.selectionStart, $bs.selectionLength)
    [PSConsoleReadLine]::Replace($bs.selectionStart, $bs.selectionLength, $repl)
    [PSConsoleReadLine]::SetCursorPosition($bs.selectionStart + $bs.selectionLength + 7)
}

##############################
# smart method completion
##############################

Remove-PSReadlineKeyHandler "tab"
Remove-PSReadlineKeyHandler "shift+tab"
Set-PSReadLineKeyHandler -Key "tab" -BriefDescription "smartCompletion" -LongDescription "insert closing parenthesis in forward completion of method" -ScriptBlock {
    param($key, $arg)
    if ($key.Modifiers -eq "Shift") {
        [PSConsoleReadLine]::TabCompletePrevious()
    }
    else {
        [PSConsoleReadLine]::TabCompleteNext()
    }
    $a = [ASTer]::new()
    if ($a.GetActiveToken().Kind -eq "LParen" -and $a.GetNextToken().Kind -ne "RParen") {
        [PSConsoleReadLine]::Insert(")")
        [PSConsoleReadLine]::BackwardChar()
        return
    }
}

##############################
# smart quotation
##############################

Set-PSReadLineKeyHandler -Key "`"","'" -BriefDescription "smartQuotation" -LongDescription "Put quotation marks and move the cursor between them or put marks around the selection" -ScriptBlock {
    param($key, $arg)
    $mark = $key.KeyChar

    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    $pos = $bs.CursorPos

    if ($bs.SelectionLength -gt 0) {
        [PSConsoleReadLine]::Replace($bs.selectionStart, $bs.selectionLength, $mark + $line.SubString($bs.selectionStart, $bs.selectionLength) + $mark)
        [PSConsoleReadLine]::SetCursorPosition($bs.selectionStart + $bs.selectionLength + 2)
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
# snippets
##############################

# Set-PSReadLineKeyHandler -Key "ctrl+k,i" -BriefDescription "insert-if-else-block" -LongDescription "insert-if-else-block" -ScriptBlock {
#     $bs = [PSBufferState]::new()
#     $pos = $bs.CursorPos
#     $indent = $bs.CursorLine.Indent
#     $filler = " " * $indent
#     $lines = @('if ($_ ) {', ($filler + '  $_'), ($filler + "}"), ($filler + "else {"), ($filler + '  $_'), ($filler + "}"))
#     [PSConsoleReadLine]::Insert($lines -join "`n")
#     [PSConsoleReadLine]::SetCursorPosition($pos + 7)
# }


Set-PSReadLineKeyHandler -Key "ctrl+k,f","ctrl+k,w" -BriefDescription "insert-alias" -LongDescription "insert-alias" -ScriptBlock {
    param ($key, $arg)
    $alias = switch ($key.KeyChar) {
        "f" { "% "; break }
        "w" { "? "; break }
    }
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + $alias)
}

Set-PSReadLineKeyHandler -Key "alt+m" -BriefDescription "measure" -LongDescription "measure" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "measure")
}

Set-PSReadLineKeyHandler -Key "alt+c" -BriefDescription "copyToClipboard" -LongDescription "copyToClipboard" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "c")
}

Set-PSReadLineKeyHandler -Key "alt+v" -BriefDescription "asVariable" -LongDescription "asVariable" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "v")
}

Set-PSReadLineKeyHandler -Key "alt+t","alt+V" -BriefDescription "teeVariable" -LongDescription "teeVariable" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "tee -Variable ")
}

Set-PSReadLineKeyHandler -Key "alt+b" -BriefDescription "bat-plain" -LongDescription "bat-plain" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "oss|bat -p")
}

Set-PSReadLineKeyHandler -Key "ctrl+k,s" -BriefDescription "insert-Select-Object" -LongDescription "insert-Select-Object" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "select -")
    [PSConsoleReadLine]::MenuComplete()
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

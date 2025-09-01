
<# ==============================

PSReadLine Configuration

                encoding: utf8bom
============================== #>

using namespace Microsoft.PowerShell
using namespace System.Management.Automation.Language

# https://github.com/PowerShell/PSReadLine/blob/dc38b451be/PSReadLine/SamplePSReadLineProfile.ps1
# you can search keychar with `[Console]::ReadKey()`

Set-PSReadlineOption -HistoryNoDuplicates `
    -PredictionSource History `
    -BellStyle None `
    -ContinuationPrompt ($Global:PSStyle.Foreground.BrightBlack + "# " + $Global:PSStyle.Reset) `
    -AddToHistoryHandler {
    param ($command)
    switch -regex ($command) {
        "SKIPHISTORY" {return $false}
        "^dsk$" {return $false}
        " -execute *$" {return $false}
        " -force *$" {return $false}
    }
    return $true
}

# Set-PSReadLineOption -colors @{
#     "Operator"          = $Global:PSStyle.Foreground.White;
# }

Set-PSReadLineKeyHandler -Key "ctrl+l" -Function ClearScreen

Set-PSReadLineKeyHandler -Key "ctrl+p","ctrl+shift+spacebar" -Function SwitchPredictionView

Set-PSReadLineKeyHandler -Key "ctrl+K" -Function DeleteLine

# search
Set-PSReadLineKeyHandler -Key "ctrl+f" -Function CharacterSearch
Set-PSReadLineKeyHandler -Key "ctrl+F" -Function CharacterSearchBackward

# accept suggestion
Set-PSReadLineKeyHandler -Key "ctrl+N" -Function AcceptSuggestion

# completion
Set-PSReadLineKeyHandler -Key "alt+i" -ScriptBlock {
    if ([PSBufferState]::IsSelecting()) {
        [PSConsoleReadLine]::DeleteChar()
    }
    [PSConsoleReadLine]::Insert("Invoke*")
}

Set-PSReadLineKeyHandler -Key "alt+-" -ScriptBlock {
    [PSConsoleReadLine]::Insert("*")
}

# custom-cd
Set-PSReadLineKeyHandler -Key "ctrl+g,d","ctrl+g,s","ctrl+g,c" -ScriptBlock {
    param($key, $arg)
    $dir = switch ($key.KeyChar) {
        <#case#> "d" { "{0}\desktop" -f $env:USERPROFILE ; break }
        <#case#> "s" { "X:\scan" ; break }
        <#case#> "c" { (Get-Clipboard | Select-Object -First 1).Replace('"', "") ; break }
    }
    if (-not (Test-Path $dir -PathType Container)) {
        $dir = $dir | Split-Path -Parent
    }
    $p = "<#SKIPHISTORY#>cd '{0}'" -f $dir
    [PSConsoleReadLine]::Insert($p)
    [PSConsoleReadLine]::AcceptLine()
}

Set-PSReadLineKeyHandler -Key "alt+p" -ScriptBlock {
    [PSConsoleReadLine]::Insert('python')
    [PSConsoleReadLine]::AcceptLine()
}

Set-PSReadLineKeyHandler -Key "alt+e" -ScriptBlock {
    [PSConsoleReadLine]::Insert('$env:')
}

# format string
Set-PSReadLineKeyHandler -Key "ctrl+k,0", "ctrl+k,1", "ctrl+k,2", "ctrl+k,3", "ctrl+k,4", "ctrl+k,5", "ctrl+k,6", "ctrl+k,7", "ctrl+k,8", "ctrl+k,9" -ScriptBlock {
    param($key, $arg)
    $str = '{' + $key.KeyChar + '}'
    [PSConsoleReadLine]::Insert($str)
}

Set-PSReadLineKeyHandler -Key "ctrl+alt+enter" -ScriptBlock {
    [PSConsoleReadLine]::Insert("return ")
}


# cursor jump
Set-PSReadLineOption -WordDelimiters ";:,.[]{}()/\|^&*-=+'`" !?@#`$%&_<>``「」（）『』『』［］、，。：；／　"
@{
    "ctrl+RightArrow"       = "ForwardWord";
    "ctrl+DownArrow"        = "ForwardWord";
    "ctrl+LeftArrow"        = "BackwardWord";
    "ctrl+UpArrow"          = "BackwardWord";
    "ctrl+shift+RightArrow" = "SelectForwardWord";
    "ctrl+shift+DownArrow"  = "SelectForwardWord";
    "ctrl+shift+LeftArrow"  = "SelectBackwardWord";
    "ctrl+shift+UpArrow"    = "SelectBackwardWord";
}.GetEnumerator() | ForEach-Object {
    Set-PSReadLineKeyHandler -Key $_.Key -Function $_.Value
}

# shell cursor jump
Set-PSReadLineKeyHandler -Key "alt+j" -Function ShellForwardWord
Set-PSReadLineKeyHandler -Key "alt+k" -Function ShellBackwardWord
Set-PSReadLineKeyHandler -Key "alt+J" -Function SelectShellForwardWord
Set-PSReadLineKeyHandler -Key "alt+K" -Function SelectShellBackwardWord

# https://github.com/pecigonzalo/Oh-My-Posh/blob/master/plugins/psreadline/psreadline.ps1
class ASTer {
    [Ast[]]$ast
    [Token[]]$tokens
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

    [Token] GetActiveToken() {
        $i = $this.GetActiveTokenIndex()
        return $this.tokens[$i]
    }

    [Token] GetLastToken() {
        return $this.tokens[$this.tokens.Count - 1]
    }

    [Token] GetPreviousToken() {
        $i = $this.GetActiveTokenIndex()
        $pos = $i - 1
        return $this.tokens[$pos]
    }

    [Token] GetNextToken() {
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
            return $this.GetActiveToken().Kind -eq [TokenKind]::Pipe
        }
        return $this.GetPreviousToken().Kind -eq [TokenKind]::Pipe
    }

    [bool] IsBeforePipe() {
        return $this.GetActiveToken().Kind -eq [TokenKind]::Pipe
    }

    ReplaceTokenByIndex([int]$index, [string]$newText) {
        $t = $this.tokens[$index]
        [PSConsoleReadLine]::Replace($t.Extent.StartOffset, ($t.Extent.EndOffset - $t.Extent.StartOffset), $newText)
    }

    ReplaceActiveToken([string]$newText) {
        $t = $this.GetActiveToken()
        [PSConsoleReadLine]::Replace($t.Extent.StartOffset, ($t.Extent.EndOffset - $t.Extent.StartOffset), $newText)
    }

}

Set-PSReadlineKeyHandler -Key "ctrl+k,d" -BriefDescription "debug-activetokenkind" -ScriptBlock {
    $a = [Aster]::new()
    $t = $a.GetActiveToken()
    Write-Host
    Write-Host $t.Kind
    Write-Host $t.Text
}

# https://github.com/pecigonzalo/Oh-My-Posh/blob/master/plugins/psreadline/psreadline.ps1
class CommandAST {
    [Ast[]]$ast
    [Token[]]$tokens
    $errors
    $cursor

    CommandAST() {
        $a = $t = $e = $c = $null
        [PSConsoleReadLine]::GetBufferState([ref]$a, [ref]$t, [ref]$e, [ref]$c)
        $this.ast = $a
        $this.tokens = $t
        $this.errors = $e
        $this.cursor = $c
    }

    [Ast[]] Listup() {
        return $this.ast.FindAll({
            $node = $args[0]
            return $node -is [CommandAst]
        }, $true)
    }

    [Ast] GetActive() {
        return $this.Listup() | Where-Object {
            return ($_.Extent.StartOffset -le $this.cursor) -and ($this.cursor -le $_.Extent.EndOffset)
        } | Select-Object -Last 1
    }

    [Ast] GetPrevious() {
        return $this.Listup() | Where-Object { $_.Extent.EndOffset -lt $this.cursor } | Select-Object -Last 1
    }

    [Ast] GetNext() {
        return $this.Listup() | Where-Object { $_.Extent.StartOffset -gt $this.cursor } | Select-Object -First 1
    }
}

# smart-accept-next-suggestion
Set-PSReadLineKeyHandler -Key "ctrl+n" -BriefDescription "smart-accept-next-suggestion" -ScriptBlock {
    $a = [ASTer]::new()
    $token = $a.GetActiveToken()
    if ($token.Kind -notin @([TokenKind]::StringExpandable, [TokenKind]::StringLiteral)) {
        $vervs = (Get-Verb).Verb
        if ($token.Text -in $vervs) {
            [PSConsoleReadLine]::Insert("-")
            return
        }
    }
    [PSConsoleReadLine]::AcceptNextSuggestionWord()
}

# smart-backward-word
Set-PSReadlineKeyHandler -Key "ctrl+backspace" -ScriptBlock {
    $a = [Aster]::new()
    $t = $a.GetActiveToken()
    $target = @(
        [TokenKind]::Function,
        [TokenKind]::Command,
        [TokenKind]::Parameter,
        [TokenKind]::EndOfInput,
        [TokenKind]::Variable
    )
    if ($t.Kind -in $target) {
        [PSConsoleReadLine]::ShellBackwardKillWord()
        return
    }
    if ($t.Text.Length -eq 1) {
        [PSConsoleReadLine]::BackwardDeleteChar()
        return
    }
    if ($a.cursor -eq $t.Extent.StartOffset) {
        [PSConsoleReadLine]::ShellBackwardKillWord()
        return
    }
    $pre = $a.GetPreviousToken()
    if ($pre.Kind -in @([TokenKind]::Command, [TokenKind]::Function)) {
        [PSConsoleReadLine]::ShellBackwardKillWord()
        return
    }
    [PSConsoleReadLine]::BackwardKillWord()
}

Set-PSReadlineKeyHandler -Key "ctrl+alt+I,t" -ScriptBlock {
    $a = [Aster]::new()
    $t = $a.GetActiveToken()
    $o = [PSCustomObject]@{
        "Kind" = $t.Kind;
        "Text" = $t.Text;
    }
    [PSConsoleReadLine]::ClearScreen()
    [System.Console]::WriteLine()
    [System.Console]::WriteLine($o)
    [PSConsoleReadLine]::SetCursorPosition($a.cursor)
}

Set-PSReadlineKeyHandler -Key "ctrl+alt+k" -ScriptBlock {
    $a = [ASTer]::new()
    $lastPipe = $a.tokens | Where-Object {$_.Kind -eq [TokenKind]::Pipe} | Where-Object {$_.Extent.EndOffset -le $a.cursor} | Select-Object -Last 1
    if ($lastPipe) {
        [PSConsoleReadLine]::SetCursorPosition($lastPipe.Extent.EndOffset - 1)
    }
}
Set-PSReadlineKeyHandler -Key "ctrl+alt+j" -ScriptBlock {
    $a = [ASTer]::new()
    $nextPipe = $a.tokens | Where-Object {$_.Kind -eq [TokenKind]::Pipe} | Where-Object {$_.Extent.StartOffset -ge $a.cursor} | Select-Object -First 1
    if ($nextPipe) {
        [PSConsoleReadLine]::SetCursorPosition($nextPipe.Extent.StartOffset + 1)
    }
}

Set-PSReadLineKeyHandler -Key "alt+l" -ScriptBlock {
    $a = [ASTer]::new()
    [PSConsoleReadLine]::Insert("|")
    if ($a.IsAfterPipe()) {
        [PSConsoleReadLine]::BackwardChar()
    }
}

Set-PSReadLineKeyHandler -Key "alt+n" -ScriptBlock {
    $a = [ASTer]::new()
    if ($a.IsAfterPipe()) {
        return
    }
    $t = $a.GetActiveToken()
    if ($t.Kind -eq [TokenKind]::EndOfInput -or $a.IsEndOfToken() -or $a.IsBeforePipe()) {
        $s = ($a.IsEndOfToken())? " -" : "-"
        [PSConsoleReadLine]::Insert($s)
        [PSConsoleReadLine]::MenuComplete()
    }
}

Set-PSReadLineKeyHandler -Key "ctrl+k,l" -BriefDescription "insert pipe before active command" -ScriptBlock {
    $ca = [CommandAST]::new()
    $activeCmd = $ca.GetActive()
    if ($activeCmd) {
        [PSConsoleReadLine]::SetCursorPosition($activeCmd.Extent.StartOffset)
        [PSConsoleReadLine]::Insert("|")
        [PSConsoleReadLine]::BackwardChar()
        return
    }
    $lastCmd = $ca.GetPrevious()
    if ($lastCmd) {
        [PSConsoleReadLine]::SetCursorPosition($lastCmd.Extent.StartOffset)
    }
    [PSConsoleReadLine]::Insert("|")
    [PSConsoleReadLine]::BackwardChar()
}

Set-PSReadLineKeyHandler -Key "ctrl+k,alt+l" -BriefDescription "insert pipe after active command" -ScriptBlock {
    $ca = [CommandAST]::new()
    $activeCmd = $ca.GetActive()
    if ($activeCmd) {
        [PSConsoleReadLine]::SetCursorPosition($activeCmd.Extent.EndOffset)
        [PSConsoleReadLine]::Insert("|")
        return
    }
    $nextCmd = $ca.GetNext()
    if ($nextCmd) {
        [PSConsoleReadLine]::SetCursorPosition($nextCmd.Extent.EndOffset)
    }
    [PSConsoleReadLine]::Insert("|")
}

class PSCursorLine {
    [string]$Text
    [string]$BeforeCursor
    [string]$AfterCursor
    [int]$Indent
    [int]$Index
    [int]$StartPos
    [int]$EndPos

    PSCursorLine([string]$line, [int]$pos) {
        $lines = $line -split "`n"
        $this.Index = ($line.Substring(0, $pos) -split "`n").Count - 1
        $this.Text = $lines[$this.Index]
        $this.Indent = $this.Text.Length - $this.Text.TrimStart().Length
        if ($this.Index -gt 0) {
            $this.StartPos = ($lines[0..($this.Index - 1)] -join "`n").Length + 1
        }
        else {
            $this.StartPos = 0
        }
        $this.EndPos = $this.StartPos + $this.Text.Length
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
    [bool]$isMultiline
    [int]$lastLineIndex

    PSBufferState() {
        $stat = $this.GetState()
        $this.Commandline = $stat.line
        $this.CursorPos = $stat.pos
        $this.CursorLine = [PSCursorLine]::new($this.Commandline, $this.CursorPos)
        $start = $length = $null
        [PSConsoleReadLine]::GetSelectionState([ref]$start, [ref]$length)
        $this.SelectionStart = $start
        $this.SelectionLength = $length
        $this.isMultiline = $this.Commandline.IndexOf("`n") -ne -1
        if ($this.isMultiline) {
            $this.lastLineIndex = ($this.Commandline -split "`n").Count - 1
        }
        else {
            $this.lastLineIndex = 0
        }
    }

    [PSCustomObject] GetState() {
        $line = $pos = $null
        [PSConsoleReadLine]::GetBufferState([ref]$line, [ref]$pos)
        return [PSCustomObject]@{
            "line" = $line;
            "pos"  = $pos;
        }
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

    [void] NewLine() {
        $line = $this.CommandLine
        $pos = $this.CursorPos
        $indent = $this.CursorLine.Indent
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

Set-PSReadLineKeyHandler -Key "Shift+Enter" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.NewLine()
}

Set-PSReadLineKeyHandler -Key "enter","ctrl+enter" -ScriptBlock {
    param($key, $arg)
    $bs = [PSBufferState]::new()
    if ($bs.isMultiline) {
        if ($key.Modifiers -eq [System.ConsoleModifiers]::Control -or $bs.CursorLine.Index -eq $bs.lastLineIndex) {
            $single = ($bs.Commandline -split "`n" | ForEach-Object {
                    $l = $_.Trim()
                    if ($l.StartsWith("#")) {
                        return ""
                    }
                    if ($l.EndsWith("{")) {
                        return $l
                    }
                    return $l + ";"
                }) -join "" -replace ";}", "}" -replace ";$", "" -replace ";+", ";"
            [PSConsoleReadLine]::AddToHistory($single)
        }
        else {
            $bs.NewLine()
            return
        }
    }
    [PSConsoleReadLine]::AcceptLine()
    if ((Get-PSReadLineOption).PredictionViewStyle -eq [PredictionViewStyle]::ListView) {
        Set-PSReadLineOption -PredictionViewStyle InlineView
    }
}

# reload
Set-PSReadLineKeyHandler -Key "alt+r", "ctrl+r" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert('<#SKIPHISTORY#> . $PROFILE')
    [PSConsoleReadLine]::AcceptLine()
}

# load clipboard
Set-PSReadLineKeyHandler -Key "ctrl+V" -ScriptBlock {
    $command = '<#SKIPHISTORY#> (gcb) -split "`n"|%{$_.Replace("`r","")}|sv CLIPPING'
    [PSConsoleReadLine]::RevertLine()
    [PSConsoleReadLine]::Insert($command)
    [PSConsoleReadLine]::AddToHistory('$CLIPPING ')
    [PSConsoleReadLine]::AcceptLine()
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

Set-PSReadLineKeyHandler -Key "alt+[" -ScriptBlock {
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

Set-PSReadLineKeyHandler -Key "ctrl+k,ctrl+j" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $curLine = $bs.CursorLine.Text
    $curLineStart = $bs.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
}
Set-PSReadLineKeyHandler -Key "ctrl+k,ctrl+k" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $curLine = $bs.CursorLine.Text
    $curLineStart = $bs.CursorLine.StartPos
    [PSConsoleReadLine]::Replace($curLineStart, $curLine.Length, $curLine+"`n"+$curLine)
    [PSConsoleReadLine]::SetCursorPosition($curLineStart)

}


# ctrl+]
Set-PSReadLineKeyHandler -Key "ctrl+Oem6" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.IndentLine()
}
# ctrl+[
Set-PSReadLineKeyHandler -Key "ctrl+Oem4" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.OutdentLine()
}

Set-PSReadLineKeyHandler -Key "ctrl+/" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $bs.ToggleLineComment()
}

Set-PSReadLineKeyHandler -Key "ctrl+|" -ScriptBlock {
    $pos = [PSBufferState]::FindMatchingPairPos()
    if ($pos -ge 0) {
        [PSConsoleReadLine]::SetCursorPosition($pos)
    }
}
Set-PSReadLineKeyHandler -Key "alt+P" -ScriptBlock {
    $pos = [PSBufferState]::FindMatchingPairPos()
    if ($pos -ge 0) {
        $start = [math]::Min($pos, $bs.CursorPos)
        $end = [math]::Max($pos, $bs.CursorPos)
        $len = $end - $start + 1
        $repl = $bs.CommandLine.Substring($start+1, $len-2)
        [PSConsoleReadLine]::Replace($start, $len, $repl)
    }
}

Set-PSReadLineKeyHandler -Key "ctrl+home" -ScriptBlock {[PSConsoleReadLine]::SetCursorPosition(0)}
Set-PSReadLineKeyHandler -Key "home" -ScriptBlock {
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

Set-PSReadLineKeyHandler -Key "ctrl+end" -ScriptBlock {
    $bs = [PSBufferState]::new()
    [PSConsoleReadLine]::SetCursorPosition($bs.Commandline.Length)
}
Set-PSReadLineKeyHandler -Key "end" -Function EndOfLine

Set-PSReadLineKeyHandler -Key "ctrl+k,v" -ScriptBlock {
    $cb = (Get-Clipboard -Raw).Trim()
    $lines = @($cb -split "\r?\n")
    $s = ($lines.Count -gt 1)? ($lines | ForEach-Object {($_ -as [string]).TrimEnd()} | Join-String -Separator "`n" -OutputPrefix "@'`n" -OutputSuffix "`n'@") : $cb
    if ([PSBufferState]::IsSelecting()) {
        [PSConsoleReadLine]::DeleteChar()
    }
    [PSConsoleReadLine]::Insert($s)
}

##############################
# redo-last-command
##############################

Set-PSReadLineKeyHandler -Key "F4" -ScriptBlock {
    [PSConsoleReadLine]::RevertLine()
    $lastCmd = ([PSConsoleReadLine]::GetHistoryItems() | Select-Object -Last 1).CommandLine
    [PSConsoleReadLine]::Insert($lastCmd)
    [PSConsoleReadLine]::AcceptLine()
}

##############################
# yank-last-argument cutomize
##############################

Set-PSReadLineKeyHandler -Key "alt+a" -ScriptBlock {
    $bs = [PSBufferState]::new()
    if ($bs.CursorLine.Index -eq 0 -and $bs.CursorLine.BeforeCursor.Trim().Length -lt 1) {
        [PSConsoleReadLine]::Insert("$")
    }
    [PSConsoleReadLine]::YankLastArg()
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

Set-PSReadLineKeyHandler -Key "ctrl+k,t" -ScriptBlock {
    $bs = [PSBufferState]::new()
    $line = $bs.CommandLine
    if ($bs.SelectionLength -lt 1) {
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
    $active = $a.GetActiveToken()
    $next = $a.GetNextToken()

    if ($active.Kind -eq [TokenKind]::LParen -and $next.Kind -ne [TokenKind]::RParen) {
        [PSConsoleReadLine]::Insert(")")
        [PSConsoleReadLine]::BackwardChar()
        return
    }
    if ($active.Kind -eq [TokenKind]::Identifier -and $next.Kind -eq [TokenKind]::RParen) {
        [PSConsoleReadLine]::DeleteChar()
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

    $a = [ASTer]::new()
    $token = $a.GetActiveToken()
    if ( ($token.Kind -eq [TokenKind]::StringLiteral -and $mark -eq '"') -or ($token.Kind -eq [TokenKind]::StringExpandable -and $mark -eq "'") ) {
        [PSConsoleReadLine]::Insert($mark)
        return
    }

    [PSConsoleReadLine]::Insert($mark + $mark)
    [PSConsoleReadLine]::SetCursorPosition($pos+1)
}


##############################
# snippets
##############################

Set-PSReadLineKeyHandler -Key "ctrl+k,f", "ctrl+k,w", "ctrl+k,alt+f","ctrl+k,alt+w", "ctrl+k,alt+F","ctrl+k,alt+W" -ScriptBlock {
    param ($key, $arg)
    $alias = switch ($key.KeyChar) {
        "f" { "% "; break }
        "w" { "? "; break }
    }
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    $suffix = ""
    if ($key.Modifiers -match "Alt") {
        if ($key.Modifiers -match "Shift") {
            $suffix = '{$_}'
        }
        else {
            $suffix = '{}'
        }
    }
    [PSConsoleReadLine]::Insert($prefix + $alias + $suffix)
    if ($suffix.Length) {
        [PSConsoleReadLine]::BackwardChar()
    }
}

Set-PSReadLineKeyHandler -Key "alt+m" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "measure")
}

Set-PSReadLineKeyHandler -Key "alt+c" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "c")
}

Set-PSReadLineKeyHandler -Key "alt+v" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "sv ")
}

Set-PSReadLineKeyHandler -Key "alt+t" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "tee -Variable ")
}

Set-PSReadLineKeyHandler -Key "alt+b" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "oss |bat -p")
}

Set-PSReadLineKeyHandler -Key "ctrl+k,s" -ScriptBlock {
    $a = [ASTer]::new()
    $prefix = ($a.IsAfterPipe())? "" : "|"
    [PSConsoleReadLine]::Insert($prefix + "select -")
    [PSConsoleReadLine]::MenuComplete()
}


##############################
# ls
##############################

Set-PSReadLineKeyHandler -Key "ctrl+b,s", "ctrl+b,e", "ctrl+b,c", "ctrl+b,S", "ctrl+b,E", "ctrl+b,C" -ScriptBlock {
    param($key, $arg)
    $opr = ($key.keychar -cin @("S", "E", "C"))? "-notlike" : "-like"
    $cmd = "ls |? Basename {0} *" -f $opr
    if ($key.keychar -ieq "c") {
        $cmd = $cmd + "*"
    }
    [PSConsoleReadLine]::Insert($cmd)
    if ($key.keychar -iin @("s", "c")) {
        [PSConsoleReadLine]::BackwardChar()
    }
}

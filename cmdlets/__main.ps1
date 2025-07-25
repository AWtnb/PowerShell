$env:DOTNET_CLI_UI_LANGUAGE = "en"

# disble progress bar
$progressPreference = "silentlyContinue"


[System.Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8")
"Output encoding: {0}" -f [System.Console]::OutputEncoding.EncodingName | Write-Host -ForegroundColor Yellow

function Reset-OutputEncodingToSJIS {
    [System.Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("shift_jis")
    "Output encoding: reset to default (shift_jis)" | Write-Host -ForegroundColor Yellow
}

function Switch-AutoHideTaskbar {
    $path = "HKCU:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StuckRects3"
    $val = (Get-ItemProperty -Path $path).Settings
    $flagIdx = 8
    # Windows11
    if ($val[$flagIdx] -eq 123) {
        $val[$flagIdx] = 122
    }
    else {
        $val[$flagIdx] = 123
    }
    Set-ItemProperty -Path $path -Name "Settings" -Value $val
    Stop-Process -ProcessName "explorer" -Force
}

function Update-Repositories {
    param (
        [parameter(Mandatory)][string]$root
    )
    $src = $env:APPDATA | Join-Path -ChildPath "git-behind-checker\update-repos.ps1"
    if (Test-Path $src) {
        & $src $root
    }
    else{
        "``{0}`` not found." -f ($src | Split-Path -Leaf) | Write-Host -ForegroundColor Red
        Write-Host "Clone from " -NoNewline
        Write-Host "https://github.com/AWtnb/git-behind-checker" -ForegroundColor Cyan -NoNewline
        Write-Host " and run ``install.ps1`` ." -NoNewline
    }
}

#################################################################
# functions arround prompt customization
#################################################################

##############################
# ime
##############################

function Get-HostProcess {
    [OutputType([System.Diagnostics.Process])]
    $p = Get-Process -Id $PID
    $i = 0
    while ($p.MainWindowHandle -eq 0) {
        if ($i -gt 10) {
            return $null
        }
        $p = $p.Parent
        $i++
    }
    return $p
}

# thanks: https://stuncloud.wordpress.com/2014/11/19/powershell_turnoff_ime_automatically/
if(-not ('Pwsh.IME' -as [type]))
{Add-Type -Namespace Pwsh -Name IME -MemberDefinition @'

[DllImport("user32.dll")]
private static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

[DllImport("imm32.dll")]
private static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);
public static int GetState(IntPtr hwnd) {
    IntPtr imeHwnd = ImmGetDefaultIMEWnd(hwnd);
    return SendMessage(imeHwnd, 0x0283, 0x0005, 0);
}

public static void SetState(IntPtr hwnd, bool state) {
    IntPtr imeHwnd = ImmGetDefaultIMEWnd(hwnd);
    SendMessage(imeHwnd, 0x0283, 0x0006, state?1:0);
}

'@ }

function Reset-ConsoleIME {
    [OutputType([System.Boolean])]
    $hostProc = Get-HostProcess
    if (-not $hostProc) {
        return $false
    }
    try {
        if ([Pwsh.IME]::GetState($hostProc.MainWindowHandle)) {
            [Pwsh.IME]::SetState($hostProc.MainWindowHandle, $false)
        }
        return $true
    }
    catch {
        return $false
    }
}

##############################
# window
##############################

if(-not ('Pwsh.Window' -as [type]))
{Add-Type -Namespace Pwsh -Name Window -MemberDefinition @'

[DllImport("user32.dll")]
private static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);
public static bool SetText(IntPtr hwnd, string text) {
    return SendMessage(hwnd, 0x000C, 0, text);
}

[DllImport("user32.dll")]
private static extern void SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
public static void Minimize(IntPtr hwnd) {
    SendMessage(hwnd, 0x0112, 0xF020, 0);
}

'@ }

function Set-ConsoleWindowTitle {
    param (
        [string]$title
    )
    $hostProc = Get-HostProcess
    if (-not $hostProc) {
        return $false
    }
    return [Pwsh.Window]::SetText($hostProc.MainWindowHandle, $title)
}

function Hide-ConsoleWindow {
    $hostProc = Get-HostProcess
    if ($hostProc -and ($env:TERM_PROGRAM -ne "vscode")) {
        [Pwsh.Window]::Minimize($hostProc.MainWindowHandle)
    }
}

function Restart-GoogleIme {
    Get-Process | Where-Object ProcessName -In @("GoogleIMEJaConverter", "GoogleIMEJaRenderer") | ForEach-Object {
        $path = $_.Path
        Stop-Process $_
        Start-Process $path
    }
    "Restarted google ime process!" | Write-Host -ForegroundColor Yellow
}

##############################
# Pseudo-voicing mark fixer
##############################

class PseudoVoicing {
    [string]$origin
    [string]$formatted
    [string]$voicables
    PseudoVoicing([string]$s) {
        $this.origin = $s
        $this.formatted = $this.origin
        $this.voicables = "かきくけこさしすせそたちつてとはひふへほカキクケコサシスセソタチツテトハヒフヘホ"
    }
    [void] FixVoicing() {
        $this.formatted = [regex]::new(".[\u309b\u3099]").Replace($this.formatted, {
            param($m)
            $c = $m.Value.Substring(0,1)
            if ($this.voicables.IndexOf($c) -lt 0) {
                return $m
            }
            if ($c -eq "う") {
                return "`u{3094}"
            }
            if ($c -eq "ウ") {
                return "`u{30f4}"
            }
            return [string]([Convert]::ToChar([Convert]::ToInt32([char]$c) + 1))
        })
    }
    [void] FixHalfVoicing() {
        $this.formatted = [regex]::new(".[\u309a\u309c]").Replace($this.formatted, {
            param($m)
            $c = $m.Value.Substring(0,1)
            if ($this.voicables.IndexOf($c) -lt 0) {
                return $m
            }
            return [string]([Convert]::ToChar([Convert]::ToInt32([char]$c) + 2))
        })
    }
}

function Rename-MacOSFile {
    param (
        [parameter(ValueFromPipeline = $true)]$inputObj
    )
    begin {
        $files = @()
    }
    process {
        $fileObj = Get-Item -LiteralPath $inputObj
        if ($fileObj.BaseName -match "\u309a|\u309b|\u309c|\u3099") {
            $files += $fileObj
        }
    }
    end {
        $fileObj  | ForEach-Object {
            "Pseudo voicing-mark on '{0}'!" -f $_.Name | Write-Host -ForegroundColor Magenta
            $ask = Read-Host "Fix? (y/n)"
            if ($ask -ne "y") {
                return
            }
            $n = [PseudoVoicing]::new($_.Name)
            $n.FixHalfVoicing()
            $n.FixVoicing()
            $_ | Rename-Item -NewName $n.formatted
            "==> Fixed!" | Write-Host
        }
    }
}


#################################################################
# prompt
#################################################################

Class Prompter {

    [string]$accentBg
    [string]$markedFg
    [string]$warningFg
    [string]$subMarkerStart
    [string]$underlineStart
    [string]$stopDeco

    Prompter() {
        $color = $this.isAdmin()? "Red" : "White"
        $this.accentBg = $Global:PSStyle.Background.PSObject.Properties[$color].Value
        $this.markedFg = $Global:PSStyle.Foreground.Black
        $this.warningFg = $Global:PSStyle.Foreground.BrightRed
        $this.subMarkerStart = $Global:PSStyle.Background.BrightBlack + $this.markedFg
        $this.underlineStart = $Global:PSStyle.Underline + $Global:PSStyle.Foreground.BrightBlack
        $this.stopDeco = $Global:PSStyle.Reset
    }

    [bool] isAdmin() {
        return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    [string] GetRepoInfo() {
        $trial = 100
        $p = (Get-Location).ProviderPath
        while ($p) {
            $trial += -1
            if ($trial -lt 0) {
                return ""
            }
            if ($p | Join-Path -ChildPath ".git" | Test-Path -PathType Container) {
                return $this.warningFg + "[.git]" + $this.stopDeco
            }
            $p = $p | Split-Path -Parent
        }
        return ""
    }

    [string] GetWd() {
        $prof = $env:USERPROFILE
        $wd = $pwd.ProviderPath
        $parent = $wd | Split-Path -Parent
        $leaf = $wd | Split-Path -Leaf
        if ($wd.StartsWith($prof)) {
            if ($wd.Length -eq $prof.length) {
                $parent = ""
                $leaf = "~"
            }
            else {
                $parent = $parent.Replace($prof, "~")
            }
        }
        $connector = ($parent.Length -lt 1 -or $parent.EndsWith("\"))? "" : "\"
        $prefix = $this.subMarkerStart + "#" + $parent + $connector + $this.stopDeco
        return $($prefix `
            + $this.accentBg `
            + $this.markedFg `
            + $leaf `
            + $this.stopDeco`
            + $this.GetRepoInfo())
    }

    [string] GetPrompt() {
        $prompt = "# "
        if (($pwd.Path | Split-Path -Leaf) -ne "Desktop") {
            return $this.warningFg + $prompt + $this.stopDeco
        }
        return $prompt
    }

    [void] Display() {
        $this.GetWd() | Write-Host
    }

}


# restart keyhac
function Restart-Keyhac {
    $procs = Get-Process -Name "keyhac" -ErrorAction SilentlyContinue
    if ($procs) {
        $path = ($procs | Select-Object -First 1).Path
        $procs | Stop-Process -Force
        Start-Process -FilePath $path
    } else {
        Start-Process -FilePath $($env:USERPROFILE | Join-Path -ChildPath "Sync\portable_app\keyhac\keyhac.exe") -PassThru
    }
}


function Set-KeyhacPriorityHigh {
    Get-Process -Name "keyhac" -ErrorAction SilentlyContinue | Where-Object {$_.PriorityClass -ne "High"} | ForEach-Object {
        $_.PriorityClass = "High"
    }
}


function prompt {
    $p = [Prompter]::New()
    $p.Display()

    if (-not (Reset-ConsoleIME)) {
        "failed to reset ime..." | Write-Host -ForegroundColor Magenta
    }

    return $p.GetPrompt()
}

#################################################################
# variable / alias / function
#################################################################

Set-Alias gd Get-Date
Set-Alias f ForEach-Object
Set-Alias w Where-Object
Set-Alias v Set-Variable
Set-Alias wh Write-Host

function Out-FileUtil {
    param (
        [string]$basename
        ,[string]$extension = "txt"
        ,[switch]$force
    )
    if ($extension.StartsWith(".")) {
        $extension = $extension.Substring(1)
    }
    if (-not $basename) {
        $basename = Get-Date -Format yyyyMMddHHmmss
    }
    $outName = ($basename.EndsWith("." + $extension))? $basename : $basename + "." + $extension
    $outPath = (Get-Location).ProviderPath | Join-Path -ChildPath $outName
    $input | Join-String -Separator "`r`n" | Out-File -FilePath $outPath -Encoding utf8NoBOM -NoClobber:$(-not $force)
}
Set-Alias of Out-FileUtil

function dsk {
    "{0}\desktop" -f $env:USERPROFILE | Set-Location
}

function d {
    Start-Process ("{0}\desktop" -f $env:USERPROFILE)
}

function sum {
    $a = @()
    if ($args.Count) {
        $args | ForEach-Object {$a += $_}
    } else {
        $input | ForEach-Object {$a += $_}
    }
    $n = 0
    $a | ForEach-Object {$n += $_}
    return $n
}

function yen {
    $reg = [regex]::new("(\d)(?=(\d{3})+$)")
    $a = @()
    if ($args.Count) {
        $args | ForEach-Object {$a += $_}
    } else {
        $input | ForEach-Object {$a += $_}
    }
    return $a | ForEach-Object {
        return $reg.Replace($_, '$1,')
    }
}

function ii. {
    Invoke-Item .
}

function noblank () {
    $input | Foreach-Object {$_ -replace "\s"} | Write-Output
}

function sieve ([switch]$net) {
    $input | Where-Object {return ($net)? ($_ -replace "\s") : $_} | Write-Output
}

function pad ([int]$width = 3, [string]$char = "0") {
    $input | ForEach-Object {($_ -as [string]).PadLeft($width, $char)} | Write-Output
}

function ml ([string]$pattern, [switch]$case, [switch]$negative){
    # ml: match line
    $reg = ($case)? [regex]::New($pattern) : [regex]::New($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($negative) {
        return @($input).Where({-not $reg.IsMatch($_)})
    }
    return @($input).Where({$reg.IsMatch($_)})
}

function ato ([string]$s) {
    return $($input | ForEach-Object {[string]$_ + $s})
}

function made ([string]$s) {
    return $($input | ForEach-Object {
        $t = $_ -as [string]
        return $t.Substring(0, $t.IndexOf($s))
    })
}

function kara ([string]$s) {
    return $($input | ForEach-Object {
        $t = $_ -as [string]
        return $t.Substring($t.IndexOf($s)+1)
    })
}

function saki ([string]$s) {
    return $($input | ForEach-Object {$s + [string]$_})
}

function sand([string]$pair = "「」") {
    if($pair.Length -eq 2) {
        $pre = $pair[0]
        $post = $pair[1]
    }
    elseif ($pair.Length -eq 1) {
        $pre = $post = $pair
    }
    elseif ($pair.Length % 2 -eq 0) {
        $l = $pair.Length / 2
        $pre = $pair.Substring(0, $l)
        $post = $pair.Substring($l)
    }
    else {
        $pre = $post = ""
    }
    return $($input | ForEach-Object {$pre + $_ + $post})
}

function reverse {
    $a = @($input)
    for ($i = $a.Count - 1; $i -ge 0; $i--) {
        $a[$i] | Write-Output
    }
}

function Invoke-Taskswitcher ([int]$waitMsec = 150) {
    Start-Sleep -Milliseconds $waitMsec
    [System.Windows.Forms.SendKeys]::SendWait("^%{Tab}")
}
function Invoke-Taskview ([int]$waitMsec = 150) {
    Start-Sleep -Milliseconds $waitMsec
    Start-Process Explorer.exe -ArgumentList @("shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}")
}
function Invoke-GodMode () {
    Start-Process Explorer.exe -ArgumentList @("shell:::{ED7BA470-8E54-465E-825C-99712043E01C}")
}

function c {
    $lines = @($input).ForEach({$_ -as [string]})
    if ($lines.Count -lt 1) {
        return
    }
    $lines | Set-Clipboard
    [System.Windows.Forms.SendKeys]::SendWait("%{Tab}")
}

function cdc {
    $clip = (Get-Clipboard | Select-Object -First 1) -replace '"'
    if (Test-Path $clip -PathType Container) {
        Set-Location $clip
    }
    else {
        "invalid-path!" | Write-Host -ForegroundColor Magenta
    }
}

function j ($i) {
    $h = [ordered]@{
        "令和" = 2018;
        "平成" = 1988;
        "昭和" = 1925;
    }
    $now = (Get-Date).Year
    $h.GetEnumerator() | ForEach-Object {
        $y = $_.Value + $i
        $ansi = ($y -gt $now)? $Global:PSStyle.Foreground.BrightBlack : $Global:PSStyle.Foreground.BrightWhite
        [PSCustomObject]@{
            "和暦" = $ansi + $_.Key + $Global:PSStyle.Reset;
            "西暦" = $ansi + $y + $Global:PSStyle.Reset;
        } | Write-Output
    }
}


function Stop-PsStyleRendering {
    $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::PlainText
}
function Start-PsStyleRendering {
    $global:PSStyle.OutputRendering = [System.Management.Automation.OutputRendering]::Ansi
}

# restart corvusskk server
function Restart-CorvusSKKServer {
    $proc = Get-Process | Where-Object {$_.Name -eq "crvskkserv"}
    if ($proc) {
        $p = $proc.Path
        $proc | Stop-Process -Force
        Start-Process $p
    }
}

# restart corvusskk
function Restart-CorvusSKK {
    param(
        [switch]$withServer
    )
    $proc = Get-Process | Where-Object {$_.Name -eq "imcrvmgr"}
    if ($proc) {
        $p = $proc.Path
        $proc | Stop-Process -Force
        Start-Process $p
    }
    if ($withServer) {
        Restart-CorvusSKKServer
    }
}

# get skk customize functions
function Get-CorvusSKKUserFunctionExamples {
    $p = $env:APPDATA | Join-Path -ChildPath "CorvusSKK\init.lua"
    if (Test-Path $p) {
        $pattern = "^- ``"
        Get-Content -Path $p | Select-String -Pattern $pattern | ForEach-Object {$_.Line -replace $pattern, "" -replace "``$", ""} | Write-Output
    }
}

# get skk lua functions in userdict
function Get-CorvusSKKLuaFunctionsInUserdict {
    $p = $env:APPDATA | Join-Path -ChildPath "CorvusSKK\userdict.txt"
    if (Test-Path $p) {
        Get-Content -Path $p | Select-String -Pattern "/\([a-z]" | ForEach-Object {
            $line = $_.Line
            return $line -split "/" | Select-Object -Skip 1 | Where-Object { $_.Trim().StartsWith("(")}
        } | Write-Output
    }
}

function Invoke-TarExtract {
    param(
        [string]$path
        ,[string]$outname
    )
    $target = Get-Item -LiteralPath $path
    if ($target.Extension -notin @(".zip", ".7z")) {
        return
    }
    if (-not $outname) {
        $outname = $target.BaseName
    }
    if (Test-Path $outname -PathType Container) {
        if ((Get-ChildItem -Path $outname).Count -gt 0) {
            "'{0}' already exists and has some contents!" -f $outname | Write-Host -ForegroundColor Magenta
            return
        }
    } else {
        New-Item -Path $outname -ItemType Directory
    }
    $params = @("-x", "-v", "-C","$outname", "-f", $target.FullName) | ForEach-Object {
        if ($_ -match "\s") {
            return '"' + $_ + '"'
        }
        return $_
    }
    Start-Process -FilePath tar.exe -ArgumentList $params -NoNewWindow -Wait
}

##############################
# update type data
##############################

Update-TypeData -TypeName "System.Object" -Force -MemberType ScriptMethod -MemberName "GetProperties" -Value {
    return $($this.PsObject.Members | Where-Object MemberType -eq noteproperty | Select-Object Name, Value)
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "CaseSensitiveEquals" -Value {
    param([string]$s)
    return [string]::Equals($this, $s, [System.StringComparison]::Ordinal)
}
Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "CaseInSensitiveEquals" -Value {
    param([string]$s)
    return [string]::Equals($this, $s, [System.StringComparison]::OrdinalIgnoreCase)
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "ToSha256" -Value {
    $bs = [System.Text.Encoding]::UTF8.GetBytes($this)
    $sha = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    $hasyBytes = $sha.ComputeHash($bs)
    $sha.Dispose()
    return $(-join ($hasyBytes | ForEach-Object {
        $_.ToString("x2")
    }))
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "WithoutSpaces" -Value {
    return $this -replace "\s", ""
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "TrimParen" -Value {
    return $this -replace "\(.+?\)|（.+?）", ""
}

Update-TypeData -TypeName "System.String" -Force -MemberType ScriptMethod -MemberName "TrimBrackat" -Value {
    return $this -replace "\[.+?\]|［.+?］", ""
}

function ConvertTo-SHA256Hash {
    param (
        [parameter(ValueFromPipeline)][string]$str
    )
    begin {}
    process {
        $str.ToSha256()
    }
    end {}
}

@("System.Double", "System.Int32") | ForEach-Object {

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "ToPaddedStr" -Value {
        param([int]$pad=2)
        return $("{0:d$($pad)}" -f [int]$this)
    }

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "ToHex" -Value {
        return $([System.Convert]::ToString($this,16))
    }

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "RoundTo" -Value {
        param ([int]$n=2)
        $digit = [math]::Pow(10, $n)
        return $([math]::Round($this * $digit) / $digit)
    }

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "FloorTo" -Value {
        param ([int]$n=2)
        $digit = [math]::Pow(10, $n)
        return $([math]::Floor($this * $digit) / $digit)
    }

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "CeilingTo" -Value {
        param ([int]$n=2)
        $digit = [math]::Pow(10, $n)
        return $([math]::Ceiling($this * $digit) / $digit)
    }

    # 13q = 9pt, 4q = 1mm

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Pt2Q" -Value {
        $q = $this * (13 / 9)
        return $q.RoundTo(1)
    }
    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Q2Pt" -Value {
        $pt = $this * (9 / 13)
        return $pt.RoundTo(1)
    }
    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Q2Mm" -Value {
        return ($this / 4).RoundTo(1)
    }
    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Mm2Q" -Value {
        return ($this * 4).RoundTo(1)
    }
    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Mm2Pt" -Value {
        $q = $this * 4
        return $q.Q2Pt()
    }
    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "Pt2Mm" -Value {
        $q = $this.Pt2Q()
        return ($q / 4).RoundTo(1)
    }

    Update-TypeData -TypeName $_ -Force -MemberType ScriptMethod -MemberName "ToCJK" -Value {
        $s = $this -as [string]
        @{
            "0" = "〇";
            "1" = "一";
            "2" = "二";
            "3" = "三";
            "4" = "四";
            "5" = "五";
            "6" = "六";
            "7" = "七";
            "8" = "八";
            "9" = "九";
        }.GetEnumerator() | ForEach-Object {
            $s = $s.Replace($_.key, $_.value)
        }
        return $s
    }
}

function  Convert-IntToCJK {
    $re = [regex]::new("\d")
    return $input | ForEach-Object {
        return $re.Replace($_, {
            param($m)
            return ($m.Value -as [int]).toCJK()
        })
    }
}

##############################
# github repo
##############################

function repo {
    ("code {0}" -f $PROFILE | Split-Path -Parent) | Invoke-Expression
}

##############################
# temp dir
##############################

function Use-TempDir {
    <#
    .NOTES
    > Use-TempDir {$pwd.Path}
    Microsoft.PowerShell.Core\FileSystem::C:\Users\~~~~~~ # includes PSProvider

    > Use-TempDir {$pwd.ProviderPath}
    C:\Users\~~~~~~ # literal path without PSProvider

    #>
    param (
        [ScriptBlock]$script
    )
    $tmp = $env:TEMP | Join-Path -ChildPath $([System.Guid]::NewGuid().Guid)
    New-Item -ItemType Directory -Path $tmp | Push-Location
    "working on tempdir: {0}" -f $tmp | Write-Host -ForegroundColor DarkBlue
    $result = $null
    try {
        $result = Invoke-Command -ScriptBlock $script
    }
    catch {
        $_.Exception.ErrorRecord | Write-Error
        $_.ScriptStackTrace | Write-Host
    }
    finally {
        Pop-Location
        $tmp | Remove-Item -Recurse
    }
    return $result
}


##############################
# highlight string
##############################

Class PsHighlight {
    [regex]$reg
    [string]$color
    PsHighlight([string]$pattern, [string]$color, [switch]$case) {
        $this.reg = ($case)? [regex]::new($pattern) : [regex]::new($pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $this.color = $color
    }
    [string]Wrap([System.Text.RegularExpressions.Match]$m) {
        return $global:PSStyle.Background.PSObject.Properties[$this.color].Value + $global:PSStyle.Foreground.Black + $m.Value + $global:PSStyle.Reset
    }
    [string]Markup([string]$s) {
        return $this.reg.Replace($s, $this.Wrap)
    }
}

function Write-StringHighLight {
    param (
        [string]$pattern
        ,[switch]$case
        ,[ValidateSet("Black","Red","Green","Yellow","Blue","Magenta","Cyan","White","BrightBlack","BrightRed","BrightGreen","BrightYellow","BrightBlue","BrightMagenta","BrightCyan","BrightWhite")][string]$color = "Yellow"
    )
    $hi = [PsHighlight]::new($pattern, $color, $case)
    foreach ($line in $input) {
        $hi.Markup($line) | Write-Output
    }
}
Set-Alias hilight Write-StringHighLight



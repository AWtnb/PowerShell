$vscode = (Get-Command "Code").Source | Split-Path -Parent | Split-Path -Parent | Join-Path -ChildPath "code.exe"
if (Test-Path $vscode) {
    $wsShell = New-Object -ComObject WScript.Shell
    $startMenu = $env:USERPROFILE | Join-Path -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
    $shortcutName = "{0}-on-vscode.lnk" -f ($PSScriptRoot | Split-Path -Leaf)
    $shortcutPath = $startMenu | Join-Path -ChildPath $shortcutName
    $shortcut = $wsShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $vscode
    $shortcut.Arguments = '"{0}"' -f $PSScriptRoot
    $shortcut.Save()
    "Created shortcut on start menu: {0}" -f $shortcutPath | Write-Host -ForegroundColor Blue
}
else {
    "Cannot find VSCode on '{0}'" -f $vscode | Write-Host -ForegroundColor Magenta
}


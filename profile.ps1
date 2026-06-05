$PSScriptRoot | Join-Path -ChildPath "cmdlets" | Get-ChildItem -Recurse -Include "*.ps1" | ForEach-Object {
    . $_.FullName
}

# https://github.com/Moeologist/scoop-completion
Import-Module "$($(Get-Item $(Get-Command scoop.ps1).Path).Directory.Parent.FullName)\modules\scoop-completion"

# https://github.com/ajeetdsouza/zoxide
Invoke-Expression (& { (zoxide init powershell | Out-String) })

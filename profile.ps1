"Loading custom cmdlets took {0:f0}ms." -f $(Measure-Command {
    $PSScriptRoot | Join-Path -ChildPath "cmdlets" | Get-ChildItem -Recurse -Include "*.ps1" | ForEach-Object {
        . $_.FullName
    }
}).TotalMilliseconds | Write-Host -ForegroundColor Cyan

# https://github.com/Moeologist/scoop-completion
Import-Module "$($(Get-Item $(Get-Command scoop.ps1).Path).Directory.Parent.FullName)\modules\scoop-completion"

# https://github.com/ajeetdsouza/zoxide
Invoke-Expression (& { (zoxide init powershell | Out-String) })

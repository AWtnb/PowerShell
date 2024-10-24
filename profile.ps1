"Loading custom cmdlets took {0:f0}ms." -f $(Measure-Command {
    $PSScriptRoot | Join-Path -ChildPath "cmdlets" | Get-ChildItem -Recurse -Include "*.ps1" | ForEach-Object {
        . $_.FullName
    }
}).TotalMilliseconds | Write-Host -ForegroundColor Cyan

New-Item $PROFILE -Force -ErrorAction SilentlyContinue
Get-Content ($PSScriptRoot | Join-Path -ChildPath "profile.ps1") | Out-File -FilePath $PROFILE -Encoding utf8
$d = "cmdlets"
New-Item -Path ($PROFILE | Split-Path -Parent | Join-Path -ChildPath $d) -Value ($PSScriptRoot | Join-Path -ChildPath $d) -ItemType Junction -Confirm -Force

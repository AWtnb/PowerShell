# My PowerShell Customize 🐚

## Install

```PowerShell
$d = "cmdlets"; New-Item $Profile -Force -ErrorAction SilentlyContinue; Get-Content .\profile.ps1 | Out-File -FilePath $PROFILE -Encoding utf8; New-Item -Path ($PROFILE | Split-Path -Parent | Join-Path -ChildPath $d) -Value ($pwd.Path | Join-Path -ChildPath $d) -ItemType Junction
```



[scoop-completion](https://github.com/Moeologist/scoop-completion)

```PowerShell
scoop bucket add extras

scoop install scoop-completion
```

[ZLocation](https://www.powershellgallery.com/packages/ZLocation/)

```PowerShell
Install-Module -Name ZLocation
```

# My PowerShell Customize 🐚

## Install

```PowerShell
New-Item $Profile -Force -ErrorAction SilentlyContinue; Get-Content .\profile.ps1 | Out-File -FilePath $PROFILE -Encoding utf8; Copy-Item -Path .\cmdlets\ -Destination ($PROFILE | Split-Path -Parent) -Recurse
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

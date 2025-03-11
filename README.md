# My PowerShell Customize 🐚

## Install

```PowerShell
$d = "cmdlets"; New-Item $PROFILE -Force -ErrorAction SilentlyContinue; Get-Content .\profile.ps1 | Out-File -FilePath $PROFILE -Encoding utf8; New-Item -Path ($PROFILE | Split-Path -Parent | Join-Path -ChildPath $d) -Value ($pwd.Path | Join-Path -ChildPath $d) -ItemType Junction
```



[scoop-completion](https://github.com/Moeologist/scoop-completion)

```
scoop bucket add extras

scoop install scoop-completion
```

[zoxide](https://github.com/ajeetdsouza/zoxide)

```
scoop install zoxide
```

[uni](https://github.com/arp242/uni) (with [chocolatey](https://chocolatey.org/install#individual))

```
choco install uni
```

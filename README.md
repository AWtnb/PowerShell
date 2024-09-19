# My PowerShell Customize 🐚

Place this repository on `$env:USERPROFILE\Documents\PowerShell` .

## Packages

[scoop-completion](https://github.com/Moeologist/scoop-completion)

```PowerShell
# add extras bucket
scoop bucket add extras

# install
scoop install scoop-completion

# enable completion in current shell, use absolute path because PowerShell Core not respect $env:PSModulePath
Import-Module "$($(Get-Item $(Get-Command scoop.ps1).Path).Directory.Parent.FullName)\modules\scoop-completion"
```

[ZLocation](https://www.powershellgallery.com/packages/ZLocation/)

```PowerShell
Install-Module -Name ZLocation
```

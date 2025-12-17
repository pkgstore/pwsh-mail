# PowerShell: Mail

Sending email using PowerShell.

## Install

```powershell
$APP = "mail"; $ORG = "pkgstore"; $PFX = "pwsh-"; $URI = "https://raw.githubusercontent.com/${ORG}/${PFX}${APP}/refs/heads/main"; $META = Invoke-RestMethod -Uri "${URI}/meta.json"; $META.install.file.ForEach({ if (-not (Test-Path "$($_.path)")) { New-Item -Path "$($_.path)" -ItemType "Directory" | Out-Null }; Invoke-WebRequest "${URI}/$($_.name)" -OutFile "$($_.path)" })
```

## Resources

- [Documentation (RU)](https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/)

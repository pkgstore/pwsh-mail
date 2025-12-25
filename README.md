# PowerShell: Mail

Sending email using PowerShell.

## Install

```powershell
$Ver = 'v0.0.0'; $App = 'mail'; Invoke-Command -ScriptBlock $([scriptblock]::Create((Invoke-WebRequest -Uri 'https://pkgstore.ru/pwsh.install.txt').Content)) -ArgumentList ($args + @($App, $Ver))
```

## Resources

- [Documentation (RU)](https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/)

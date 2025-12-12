<#PSScriptInfo
.VERSION      0.1.0
.GUID         3f9da429-27eb-46de-a72c-1b0ba149a30e
.AUTHOR       Kai Kimera
.AUTHOREMAIL  mail@kaikim.ru
.TAGS         windows server mail
.LICENSEURI   https://choosealicense.com/licenses/mit/
.PROJECTURI   https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/
#>

#Requires -Version 7.2

<#
.SYNOPSIS
Sends an email notification using SMTP.

.DESCRIPTION

.EXAMPLE
.\app.mail.ps1 -Subject 'Example' -Body 'Hello world!' -From 'mail@example.com' -To 'mail@example.org'

.EXAMPLE
.\app.mail.ps1 -Subject 'Example' -Body 'Hello world!' -From 'mail@example.com' -To 'mail@example.org' -Attachment 'C:\file.01.txt', 'C:\file.02.txt'

.LINK
https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/
#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [string]$Hostname = ([System.Net.Dns]::GetHostEntry([System.Environment]::MachineName).HostName),
  [Parameter(Mandatory)][string]$Subject,
  [Parameter(Mandatory)][string]$Body,
  [Parameter(Mandatory)][string]$From,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$To,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Cc,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Bcc,
  [SupportsWildcards()][string[]]$File,
  [ValidateSet('Low', 'Normal', 'High')][string]$Priority = 'Normal',
  [switch]$Wildcard,
  [switch]$Rename,
  [switch]$Remove,
  [switch]$HTML,
  [switch]$SSL,
  [switch]$BypassCertValid
)

$CFG = ((Get-Item "${PSCommandPath}").Basename + '.ini');
$P = (Get-Content -Path "${PSScriptRoot}\${CFG}" | ConvertFrom-StringData)
$LOG = "${PSScriptRoot}\log.mail.txt"
$UUID = (Get-CimInstance 'Win32_ComputerSystemProduct' | Select-Object -ExpandProperty 'UUID')
$HID = (-join ($Hostname, ':', $UUID).ToUpper())
$DATE = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
$NL = [Environment]::NewLine

if ($Wildcard) {
  $File = (Resolve-Path "${File}" | Select-Object -ExpandProperty 'Path'); if ($null -eq $File ) { exit }
} else {
  $File.ForEach({ if (-not (Test-Path -LiteralPath "${_}" -PathType 'Leaf')) { exit } })
}

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function Write-Sign {
  $Sign = switch ( $true ) {
    $HTML {
      -join (
        '<br><br>-- <ul>',
        "<li><pre><code>#ID:${HID}</code></pre></li>",
        "<li><pre><code>#DATE:${DATE}</code></pre></li>",
        '</ul>'
      )
    }
    default {
      -join (
        "${NL}${NL}-- ",
        "${NL}#ID:${HID}",
        "${NL}#DATE:${DATE}"
      )
    }
  }

  return $Sign
}

function Update-File {
  $File.ForEach({
    if ($Rename) { Move-Item -LiteralPath "${_}" -Destination "${_}.attach" -Force }
    if ($Remove) { Remove-Item -LiteralPath "${_}" -Force }
  })
}

function Send-Mail {
  try {
    $Mail = (New-Object System.Net.Mail.MailMessage)
    $Mail.Subject = $Subject
    $Mail.Body = (-join ($Body, $(Write-Sign)))
    $Mail.From = $From
    $Mail.Priority = $Priority
    $Mail.IsBodyHtml = $HTML
    $To.ForEach({ $Mail.To.Add($_) })
    $Cc.ForEach({ $Mail.CC.Add($_) })
    $Bcc.ForEach({ $Mail.BCC.Add($_) })
    $File.ForEach({ $Mail.Attachments.Add((New-Object System.Net.Mail.Attachment($_))) })

    if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }
    $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
    $SmtpClient.EnableSsl = $SSL
    $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
    $SmtpClient.Send($Mail)
    Write-Host "Email sent successfully!${NL}" -ForegroundColor 'Green'
  } catch {
    Write-Error "ERROR: $($_.Exception.Message)"
  } finally {
    $Mail.Dispose()
    $SmtpClient.Dispose()
    Update-File
  }
}

function Start-Script() {
  Start-Transcript -Path "${LOG}"
  Send-Mail
  Stop-Transcript
}; Start-Script

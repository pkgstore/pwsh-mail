<#PSScriptInfo
.VERSION      0.1.0
.GUID         3f9da429-27eb-46de-a72c-1b0ba149a30e
.AUTHOR       Kai Kimera
.AUTHOREMAIL  mail@kai.kim
.TAGS         windows server mail
.LICENSEURI   https://choosealicense.com/licenses/mit/
.PROJECTURI   https://libsys.ru/ru/
#>

<#
.SYNOPSIS
Sends an email notification using SMTP.

.DESCRIPTION

.PARAMETER Subject
The subject of the email.

.PARAMETER Body
The body of the email. Can be plain text or HTML based on the 'HTML' flag.

.PARAMETER From
The email address of the sender.

.PARAMETER To
An array of recipient email addresses.

.PARAMETER Cc
An optional array of CC recipient email addresses.

.PARAMETER Bcc
An optional array of BCC recipient email addresses.

.PARAMETER Attachment
An optional file path for an email attachment.

.PARAMETER Priority

.PARAMETER HTML

.PARAMETER SSL

.PARAMETER BypassCertValid

.EXAMPLE
.\app.mail.ps1 -Subject 'Example' -Body 'Hello world!' -From 'mail@example.com' -To 'mail@example.org'

.EXAMPLE
.\app.mail.ps1 -Subject 'Example' -Body 'Hello world!' -From 'mail@example.com' -To 'mail@example.org' -Attachment 'C:\file.01.txt', 'C:\file.02.txt'

.NOTES
This function requires appropriate network permissions to access the SMTP server.

.LINK
https://libsys.ru/
#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [Parameter(Mandatory)][Alias('S','Subj')][string]$Subject,
  [Parameter(Mandatory)][Alias('B','Text')][string]$Body,
  [Parameter(Mandatory)][Alias('F')][string]$From,
  [Parameter(Mandatory)][Alias('T')][string[]]$To,
  [Alias('C','Copy')][string[]]$Cc,
  [Alias('BC','HideCopy')][string[]]$Bcc,
  [Alias('A','File')][string[]]$Attachment,
  [ValidateSet('Low','Normal','High')][Alias('P')][string]$Priority = 'Normal',
  [Alias('H','Host')][string]$Hostname = ([System.Net.Dns]::GetHostEntry($env:ComputerName).HostName),
  [switch]$HTML,
  [switch]$SSL,
  [switch]$BypassCertValid
)

$S = ((Get-Item "${PSCommandPath}").Basename + '.ini')
$P = (Get-Content -Path "${PSScriptRoot}\${S}" | ConvertFrom-StringData)
$UUID = (Get-CimInstance 'Win32_ComputerSystemProduct' | Select-Object -ExpandProperty 'UUID')
$HID = ((${Hostname} + ':' + ${UUID}).ToUpper())
$DATE = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
$NL = [Environment]::NewLine

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function Write-Sign {
  $Sign = switch ( $true ) {
    $HTML {
      "<br><br>-- <ul>" +
      "<li><pre><code>#ID:${HID}</code></pre></li>" +
      "<li><pre><code>#DATE:${DATE}</code></pre></li>" +
      "</ul>"
    }
    default {
      "${NL}${NL}-- " +
      "${NL}#ID:${HID}" +
      "${NL}#DATE:${DATE}"
    }
  }

  return $Sign
}

function Write-Mail {
  $Mail = (New-Object System.Net.Mail.MailMessage)
  $Mail.Subject = $Subject
  $Mail.Body = $Body + $(Write-Sign)
  $Mail.From = $From
  $Mail.Priority = $Priority
  $Mail.IsBodyHtml = $HTML

  $To.ForEach({ $Mail.To.Add($_) })
  $Cc.ForEach({ $Mail.CC.Add($_) })
  $Bcc.ForEach({ $Mail.BCC.Add($_) })
  $Attachment.ForEach({ $Mail.Attachments.Add($(New-Object System.Net.Mail.Attachment($_))) })

  return $Mail
}

function ConvertTo-String ($data) {
  $data = ($data | Join-String -SingleQuote -Separator ', ')
  return $data
}

function Start-Smtp {
  try {
    if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }
    $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
    $SmtpClient.EnableSsl = $SSL
    $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
    $SmtpClient.Send($(Write-Mail))
    Write-Host "Email $(ConvertTo-String (Write-Mail).Subject) from $(ConvertTo-String (Write-Mail).From) to $(ConvertTo-String (Write-Mail).To) sent successfully!"
  } catch {
    Write-Error "ERROR: $($_.Exception.Message)"
  }
}

function Start-Script() {
  Start-Smtp
}; Start-Script

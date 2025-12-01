<#PSScriptInfo
.VERSION      0.1.0
.GUID         3f9da429-27eb-46de-a72c-1b0ba149a30e
.AUTHOR       Kai Kimera
.AUTHOREMAIL  mail@kai.kim
.TAGS         windows server mail
.LICENSEURI   https://choosealicense.com/licenses/mit/
.PROJECTURI   https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/
#>

#Requires -Version 7.2

<#
.SYNOPSIS
Sends an email notification using SMTP.

.DESCRIPTION

.LINK
https://libsys.ru/ru/2025/12/1f77872e-d835-510b-9dc0-99ac3b4abadf/
#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [Parameter(Mandatory)]
  [Alias('S', 'Subj')]
  [string]$Subject,

  [Parameter(Mandatory)]
  [Alias('B', 'Text')]
  [string]$Body,

  [Parameter(Mandatory)]
  [Alias('F')]
  [string]$From,

  [Parameter(Mandatory)]
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')]
  [Alias('T')]
  [string[]]$To,

  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')]
  [Alias('C', 'Copy')]
  [string[]]$Cc,

  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')]
  [Alias('BC', 'HideCopy')]
  [string[]]$Bcc,

  [Alias('A','File')]
  [string[]]$Attachment,

  [ValidateSet('Low', 'Normal', 'High')]
  [Alias('P')]
  [string]$Priority = 'Normal',

  [Alias('H', 'Host')]
  [string]$Hostname = ([System.Net.Dns]::GetHostEntry($env:ComputerName).HostName),

  [switch]$HTML,
  [switch]$SSL,
  [switch]$BypassCertValid
)

$CFG = ((Get-Item "${PSCommandPath}").Basename + '.ini');
$P = (Get-Content -Path "${PSScriptRoot}\${CFG}" | ConvertFrom-StringData)
$LOG = "${PSScriptRoot}\log.mail.txt"
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

  $Attachment.ForEach({
    if (Test-Path -Path "${_}" -PathType 'Leaf') {
      $Mail.Attachments.Add($(New-Object System.Net.Mail.Attachment($_)))
    }
  })

  return $Mail
}

function Write-Status {
    $Data = @(
      [PSCustomObject]@{Name='Subject'; Value=(Write-Mail).Subject}
      [PSCustomObject]@{Name='From'; Value=(Write-Mail).From}
      [PSCustomObject]@{Name='To'; Value=(Write-Mail).To}
      [PSCustomObject]@{Name='CC'; Value=(Write-Mail).CC}
      [PSCustomObject]@{Name='BCC'; Value=(Write-Mail).BCC}
      [PSCustomObject]@{Name='Priority'; Value=(Write-Mail).Priority}
      [PSCustomObject]@{Name='HTML'; Value=(Write-Mail).IsBodyHtml}
      [PSCustomObject]@{Name='Attachment'; Value=(Write-Mail).Attachments.Name}
    ); $Data | Select-Object @{
      Name='Name'; Expression={$_.Name.PadRight(11)}
    }, @{
      Name='Value'; Expression={$_.Value | Join-String -Separator ', '}
    } | ForEach-Object { Write-Host "$($_.Name): $($_.Value)" -ForegroundColor 'Yellow' }
}

function Start-Smtp {
  try {
    if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }
    $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
    $SmtpClient.EnableSsl = $SSL
    $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
    $SmtpClient.Send($(Write-Mail))
    Write-Host "Email sent successfully!${NL}" -ForegroundColor 'Green' && $(Write-Status)
  } catch {
    Write-Error "ERROR: $($_.Exception.Message)"
  }
}

function Start-Script() {
  Start-Transcript -Path "${LOG}"
  Start-Smtp
  Stop-Transcript
}; Start-Script

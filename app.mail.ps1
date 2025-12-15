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
  [string]$Subject = (Get-Content -Path "${PSScriptRoot}\lib.mail.subject" -Encoding 'UTF8'),
  [string]$Body = (Get-Content -Path "${PSScriptRoot}\lib.mail.body" -Encoding 'UTF8' -Raw),
  [string]$Sign = (Get-Content -Path "${PSScriptRoot}\lib.mail.sign" -Encoding 'UTF8' -Raw),
  [Parameter(Mandatory)][string]$From,
  [Parameter(Mandatory)][ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$To,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Cc,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Bcc,
  [SupportsWildcards()][string[]]$File,
  [ValidateSet('Low', 'Normal', 'High')][string]$Priority = 'Normal',
  [string]$Storage = 'C:\Storage\Email',
  [int]$Count = 4,
  [switch]$Wildcard,
  [switch]$FileMove,
  [switch]$FileRemove,
  [switch]$FileList,
  [switch]$HTML,
  [switch]$SSL,
  [switch]$NoSign,
  [switch]$NoMeta,
  [switch]$BypassCertValid,
  [switch]$DateTime
)

$CFG = ((Get-Item "${PSCommandPath}").Basename + '.ini')
$P = (Get-Content -Path "${PSScriptRoot}\${CFG}" | ConvertFrom-StringData)
$LOG = "${PSScriptRoot}\log.mail.txt"
$UUID = (Get-CimInstance 'Win32_ComputerSystemProduct' | Select-Object -ExpandProperty 'UUID')
$HID = (-join ($Hostname, ':', $UUID).ToUpper())
$DATE = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
$NL = [Environment]::NewLine
$TS = $DateTime ? (Get-Date -UFormat '%F.%H-%M-%S' -AsUTC) : (Get-Date -UFormat '%s')

if ($Wildcard) {
  $File = (Resolve-Path "${File}" | Select-Object -ExpandProperty 'Path'); if ($null -eq $File ) { exit }
} else {
  $File.ForEach({ if (-not (Test-Path -LiteralPath "${_}" -PathType 'Leaf')) { exit } })
}

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function New-Storage([string]$Path) {
  if (-not (Test-Path -LiteralPath "${Path}" -PathType 'Container')) {
    New-Item -Path "${Path}" -ItemType 'Directory' | Out-Null
  }
}

function Remove-Storage([string]$Path = $Storage, [int]$Count = $Count) {
  if (Test-Path -LiteralPath "${Path}" -PathType 'Container') {
    Get-ChildItem -Path "${Path}" -Directory | Sort-Object 'CreationTime' -Descending | Select-Object -Skip $Count
      | Remove-Item -Recurse -Force
  }
}

function Move-File([string]$Path) {
  $Dir = (Join-Path -Path "${Storage}" -ChildPath "${TS}")
  New-Storage -Path "${Dir}" && Move-Item -LiteralPath "${Path}" -Destination "${Dir}"
}

function Remove-File([string]$Path) {
  Remove-Item -LiteralPath "${Path}"
}

function Write-Sep {
  $Sign = switch ( $true ) {
    $HTML   { -join ('<br><br>', '<hr style="border:none;border-top:1px solid #cccccc;width:100%;">') }
    default { -join ("${NL}${NL}-- ", "${NL}") }
  }

  return $Sign
}

function Write-Sign {
  if ($NoSign) { return }

  $Sign = switch ( $true ) {
    $HTML   { -join ("${Sign}", '<br>') }
    default { -join ("${Sign}", "${NL}") }
  }

  return $Sign
}

function Write-Meta {
  if ($NoMeta) { return }

  $Meta = switch ( $true ) {
    $HTML   { -join ('<ul>', "<li><code>#ID:${HID}</code></li>", "<li><code>#DATE:${DATE}</code></li>", '</ul>') }
    default { -join ("#ID:${HID}${NL}", "#DATE:${DATE}${NL}") }
  }

  return $Meta
}

function Write-FileList {
  if (-not $FileList) { return }

  $FileList = switch ( $true ) {
    $HTML   { -join ('<br><br><ul>', ($File.ForEach({ "<li><code>${_}</code></li>" }) | Join-String), '</ul>') }
    default { -join ("${NL}${NL}", ($File.ForEach({ "${_}" }) | Join-String -Separator "${NL}")) }
  }

  return $FileList
}

function Update-File {
  Remove-Storage
  $File.ForEach({
    if ($FileMove) { Move-File -Path "${_}" }
    if ($FileRemove) { Remove-File -Path "${_}" }
  })
}

function Send-Mail {
  try {
    $Mail = (New-Object System.Net.Mail.MailMessage)
    $Mail.Subject = $Subject
    $Mail.Body = (-join ($Body, $(Write-FileList), $(Write-Sep), $(Write-Sign), $(Write-Meta)))
    $Mail.BodyEncoding= $([System.Text.Encoding]::UTF8)
    $Mail.From = $From
    $Mail.Priority = $Priority
    $Mail.IsBodyHtml = $HTML
    $To.ForEach({ $Mail.To.Add($_) })
    $Cc.ForEach({ $Mail.CC.Add($_) })
    $Bcc.ForEach({ $Mail.BCC.Add($_) })

    if (-not $FileList) {
      $File.ForEach({ $Mail.Attachments.Add((New-Object System.Net.Mail.Attachment($_))) })
    }

    if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }
    $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
    $SmtpClient.EnableSsl = $SSL
    $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
    $SmtpClient.Send($Mail) && Write-Host "Email sent successfully!${NL}" -ForegroundColor 'Green'

    $Info = @(
      [PSCustomObject]@{Name='Subject'; Value=$Mail.Subject}
      [PSCustomObject]@{Name='From'; Value=$Mail.From}
      [PSCustomObject]@{Name='To'; Value=$Mail.To}
      [PSCustomObject]@{Name='CC'; Value=$Mail.CC}
      [PSCustomObject]@{Name='BCC'; Value=$Mail.BCC}
      [PSCustomObject]@{Name='Priority'; Value=$Mail.Priority}
      [PSCustomObject]@{Name='HTML'; Value=$Mail.IsBodyHtml}
      [PSCustomObject]@{Name='Attachments'; Value=$Mail.Attachments.Name}
    ); $Info | Select-Object @{
      Name='Name'; Expression={$_.Name.PadRight(12)}
    }, @{
      Name='Value'; Expression={$_.Value | Join-String -Separator ', '}
    } | ForEach-Object { Write-Host "$($_.Name): $($_.Value)" -ForegroundColor 'Yellow' }
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

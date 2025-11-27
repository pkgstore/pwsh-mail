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

.DESCRIPTION

.EXAMPLE
.\app.mail.ps1 [-SSL]

.LINK

#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [Parameter(Mandatory)][Alias('S')][string]$Subject,
  [Parameter(Mandatory)][Alias('B')][string]$Body,
  [Parameter(Mandatory)][Alias('F')][string]$From,
  [Parameter(Mandatory)][Alias('T')][string[]]$To,
  [string[]]$Cc,
  [string[]]$Bcc,
  [Alias('A')][string[]]$Attachment,
  [ValidateSet('Low', 'Normal', 'High')][Alias('P')][string]$Priority = 'Normal',
  [Alias('H')][string]$Hostname = ([System.Net.Dns]::GetHostEntry($env:ComputerName).HostName),
  [switch]$HTML = $false,
  [switch]$SSL = $false
)

$S = ((Get-Item "${PSCommandPath}").Basename + '.ini')
$P = (Get-Content -Path "${PSScriptRoot}\${S}" | ConvertFrom-StringData)
$UUID = (Get-CimInstance 'Win32_ComputerSystemProduct' | Select-Object -ExpandProperty 'UUID')
$HID = ((${Hostname} + ':' + ${UUID}).ToUpper())
$DATE = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function Start-Smtp {
  param(
    [Alias('S')][string]$Subject,
    [Alias('B')][string]$Body
  )

  $MailMessage = (New-Object System.Net.Mail.MailMessage)
  $MailMessage.Subject = $Subject
  $MailMessage.Body = $Body
  $MailMessage.From = $From
  $MailMessage.IsBodyHtml = $HTML

  $To.ForEach({ $MailMessage.To.Add($_) })
  $Cc.ForEach({ $MailMessage.CC.Add($_) })
  $Bcc.ForEach({ $MailMessage.BCC.Add($_) })
  $Attachment.ForEach({ $MailMessage.Attachments.Add($(New-Object System.Net.Mail.Attachment($_))) })

  $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
  $SmtpClient.EnableSsl = $SSL
  $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
  $SmtpClient.Send($MailMessage)
}

function Send-Msg() {
  $Subject = "${Subject}"
  $Body = @"
${Body}

--
#ID:${HID}
#DATE:${DATE}
"@

  Start-Smtp -Subject "${Subject}" -Body "${Body}"
}

function Start-Script() {
  Send-Msg
}; Start-Script

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
  [Alias('CC')][string[]]$CC,
  [Alias('BCC')][string[]]$BCC,
  [Alias('A')][string[]]$Attachment,
  [ValidateSet('Low', 'Normal', 'High')][Alias('P')][string]$Priority = 'Normal',
  [Alias('H')][string]$Hostname = ([System.Net.Dns]::GetHostByName([string]'localhost').HostName),
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
  $MailMessage.IsBodyHtml = $False

  if ($To.Count -gt 0) { foreach ($Person in $To) { $MailMessage.To.Add($Person) } }
  if ($CC.Count -gt 0) { foreach ($Person in $CC) { $MailMessage.CC.Add($Person) } }
  if ($BCC.Count -gt 0) { foreach ($Person in $BCC) { $MailMessage.BCC.Add($Person) } }

  if ($Attachment.Count -gt 0) {
      foreach ($File in $Attachment) {
          $Extension = (((Get-ChildItem -Path $File.FilePath).extension).ToLower())
          switch ($Extension) {
            '.gif'  { $ContentType = 'Image/gif' }
            '.jpg'  { $ContentType = 'Image/jpeg' }
            '.jpeg' { $ContentType = 'Image/jpeg' }
            '.png'  { $ContentType = 'Image/png' }
            '.csv'  { $ContentType = 'text/csv' }
            '.txt'  { $ContentType = 'text/plain' }
          }
          $Attachment = @()
          $Attachment += (New-Object System.Net.Mail.Attachment($File.FilePath, $ContentType))
          if ($null -ne $File.ContentID) { $Attachment[-1].ContentID = $File.ContentID }
          if ($ContentType.Substring(0,4) -eq 'text') {
            $Attachment[-1].ContentDisposition.FileName = ((Get-ChildItem -Path $File.FilePath).Name)
          }
          $MailMessage.Attachments.Add($Attachment[-1])
      }
  }

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
}; Start-Script

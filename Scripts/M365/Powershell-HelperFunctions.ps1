<#
.SYNOPSIS

Configuration as Code - Azure DevOps Scafold and Pipeline
Copyright (C) 2021  Sjoerd de Valk

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License version 3 as published by
the Free Software Foundation.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
All helper functions for use by the default Powershell 5.1+ implementation are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\Powershell-HelperFunctions.ps1
#>

function New-AppSelfsignedCertificate([object]$app) {
  # Create the self signed cert
  $certPfxPrefix = $global:jsonenvironmentMain.customerPrefix.toLower()
  $currentDate = Get-Date
  $endDate = $currentDate.AddYears(2)
  $notAfter = $endDate.AddYears(2)
  $pwdplain = [Guid]::NewGuid()
  $thumb = (New-SelfSignedCertificate -CertStoreLocation cert:\currentuser\my -DnsName "local.foo.$certPfxPrefix" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter).Thumbprint
  $pwdsec = ConvertTo-SecureString -String $pwdplain -Force -AsPlainText
  $exportPfx = Export-PfxCertificate -cert "cert:\currentuser\my\$thumb" -FilePath "$($O365ScriptRootLevel)$certPfxPrefix-$($app.AppId)-$Environment.pfx" -Password $pwdsec

  # Load the certificate
  $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate("$((Resolve-Path -Path $($O365ScriptRootLevel)).Path)\$certPfxPrefix-$($app.AppId)-$Environment.pfx", $pwdsec)
  $keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())

  # Add the Azure Active Directory Application Certifcate
  $secret = New-AzureADApplicationKeyCredential -ObjectId $app.ObjectId -CustomKeyIdentifier "MySecret" -StartDate $currentDate -EndDate $endDate -Type AsymmetricX509Cert -Usage Verify -Value $keyValue
  return @{"CertThumb" = $thumb; "PfxPassword" = $pwdplain}
}


function Send-Report([string[]]$mailto, [string]$subject, [string]$body, [string]$file) {
  try {
    Write-Host "Sending email"
    $params = @{}
    $params['From'] = "$($global:dstCred.UserName)"
    $params['To'] = $mailto
    $params['Body'] = $body
    $params['Subject'] = $subject
    if ($file.Length -gt 0) {
      $params['Attachments'] = $file
    }
    Send-MailMessage `
      @params `
      -BodyAsHtml `
      -Credential $global:dstCred `
      -UseSSl `
      -Port '587' `
      -SmtpServer 'smtp.office365.com' `
      -EA Stop
  }
  catch {
    Write-Host "Error sending email: $($Error[0].ToString())" -ForegroundColor "Red"
    return
  }
}

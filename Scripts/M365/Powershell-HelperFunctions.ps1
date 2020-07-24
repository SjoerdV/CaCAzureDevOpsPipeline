<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
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

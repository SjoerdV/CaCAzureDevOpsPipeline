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
All helper functions for the Exchange Online module are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . ./Az-HelperFunctions.ps1
#>


function Connect-Az([object]$properties) {
  try {
    switch ($properties.AuthSchemeType) {
      "Cred" {
        Write-Host "Connecting to Microsoft Az CLI with a user account '$($global:dstcred.UserName)'" -ForegroundColor "Green"
        $disconnectaz = az logout | Out-Null
        $Confirm = Read-Host "Is your account '$($global:dstCred.UserName)' MFA enabled? (Y/N)?"
        if($Confirm -match "[y]") {
          $connectaz = az login --allow-no-subscriptions -t $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain
        }
        else {
          $connectaz = az login --allow-no-subscriptions -t $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -u $global:dstCred.UserName -p $global:dstCred.GetNetworkCredential().Password
        }
      }
      "Thumb" {
        if ($global:dstGraphCredCertThumb) {
          Write-Host "Connecting to Microsoft Az CLI as App with a thumbprint is not supported at this time." -ForegroundColor "Green"
        }
        else {
          Write-Host "No valid 'Certificate Thumbprint' detected. Exiting..." -ForegroundColor "Red"
          exit
        }
      }
      "PfxFile" {
        if ($global:dstGraphCredCertPfxFilePath) {
          Write-Host "Connecting to Microsoft Az CLI as $($global:dstGraphCred.UserName) with Certificate Pfx File" -ForegroundColor "Green"
          # Set vars and check or download OpenSSL
          if ($IsWindows) {
            $pfxPath = "$($global:dstAzGraphCredCertPfxFilePath)"
            $rootPath = $O365ScriptRootLevel
          }
          else {
            $pfxPath = "$(($global:dstAzGraphCredCertPfxFilePath -replace '\\','/') -replace ':','')"
            $rootPath = "$(($O365ScriptRootLevel -replace '\\','/') -replace ':','')"
          }
          try {
            $versionInfo = (openssl version)
          }
          catch {
            $versionInfo = $null
          }
          if (!$versionInfo) {
            Write-Host "You need to install OpenSSL first. Either build from source: https://github.com/openssl/openssl#build-and-install or through you favorite package manager. Exiting..."
            exit
          }
          # Check or create PEM file from PFX in ASCII encoding
          $cert = (openssl pkcs12 -in "$($pfxPath)" -clcerts -nokeys -passin pass:$global:dstGraphCredCertPfxPasswordPlain)
          [array]$newcert = @()
          foreach ($line in $cert) {
            if ($line -notmatch '^(\s|Bag Attributes|subject|issuer).*' -and $line -ne '') {
              $newcert += $line
            }
          }
          $key = (openssl pkcs12 -in "$($pfxPath)" -nocerts -nodes -passin pass:$global:dstGraphCredCertPfxPasswordPlain)
          [array]$newkey = @()
          foreach ($line in $key) {
            if ($line -notmatch '^(\s|Bag Attributes|Key Attributes).*' -and $line -ne '') {
              $newkey += $line
            }
          }
          ## Merge Cert and Key and Export to a PEM file
          $totalkey = $newkey + $newcert
          $totalkey | Out-File -FilePath "$($rootPath)tempcert.pem" -Encoding ascii -Force

          # Connect to Microsoft Az CLI
          $connect = az login --service-principal --username "$($global:dstGraphCred.UserName)" --tenant "$($global:jsonenvironmentMisc.AzureADTenantId)" --password "$($rootPath)tempcert.pem"

          # Cleanup
          $del = Remove-Item -Path "$($rootPath)*" -Include *.pem -Force -ErrorAction SilentlyContinue
        }
        else {
          Write-Host "No valid 'Certificate PFX File' detected. Exiting..." -ForegroundColor "Red"
          exit
        }
      }
      default {
        Write-Host "No Connection type was specified. Exiting..." -ForegroundColor "Red"
        exit
      }
    }
  }
  catch {
    Write-Host "Connecting to Microsoft Az CLI failed: $($Error[0].ToString())" -ForegroundColor "Red"
    exit
  }
}


function New-AppSelfsignedCertificate([object]$app) {
  # Create the self signed cert
  $certPfxPrefix = $global:jsonenvironmentMain.customerPrefix.toLower()
  $pwdplain = [Guid]::NewGuid()
  if ($IsWindows) {
    $currentDate = Get-Date
    $endDate = $currentDate.AddYears(2)
    $notAfter = $endDate.AddYears(2)
    $thumb = (New-SelfSignedCertificate -CertStoreLocation cert:\currentuser\my -DnsName "$($app.displayName) - $($Environment)" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter).Thumbprint
    $pwdsec = ConvertTo-SecureString -String $pwdplain -Force -AsPlainText
    $exportPfx = Export-PfxCertificate -cert "cert:\currentuser\my\$thumb" -FilePath "$($O365ScriptRootLevel)$certPfxPrefix-$($app.AppId)-$($Environment).pfx" -Password $pwdsec

    # Load the certificate
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate("$((Resolve-Path -Path $($O365ScriptRootLevel)).Path)$($global:folderseparator)$certPfxPrefix-$($app.appId)-$Environment.pfx", $pwdsec)
    $keyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())
  }
  else {
    $pemCertFile = ("$($O365ScriptRootLevel)$certPfxPrefix-$($app.AppId)-$($Environment)_cert.pem" -replace "\\","/") -replace ":",""
    $pemKeyFile = ("$($O365ScriptRootLevel)$certPfxPrefix-$($app.AppId)-$($Environment)_key.pem" -replace "\\","/") -replace ":",""
    $pfxFile = ("$($O365ScriptRootLevel)$certPfxPrefix-$($app.AppId)-$($Environment).pfx" -replace "\\","/") -replace ":",""
    $requestgen = openssl req -x509 -newkey rsa:4096 -keyout "$pemKeyFile" -out "$pemCertFile" -sha256 -days 730 -subj "/CN=$($app.displayName) - $($Environment)" -passout pass:$pwdplain
    $pfxgen = openssl pkcs12 -export -out "$pfxFile" -inkey "$pemKeyFile" -in "$pemCertFile" -passin pass:$pwdplain -passout pass:$pwdplain
    $thumb = ((openssl x509 -in "$pemCertFile" -noout -fingerprint).split('=')[1] -replace ':').ToUpper()

    # Load the certificate
    $certArray = openssl x509 -in "$pemCertFile" -passin pass:$pwdplain
    $startLineIndex = 0
    $finalLineIndex = $certArray.Count -1
    $currentLineIndex = 0
    $keyValue = ""
    foreach ($line in $certArray) {
      if ($currentLineIndex -ne $startLineIndex -and $currentLineIndex -ne $finalLineIndex) {
        $keyValue += $line
      }
      $currentLineIndex++
    }
    # Cleanup
    $del = Remove-Item -Path "$(($O365ScriptRootLevel -replace '\\','/') -replace ':','')*" -Include *.pem -Force -ErrorAction SilentlyContinue
  }

  # Add the Azure Active Directory Application Certifcate
  # -------------------------------------------------------------------------------
  # Upload a .cer file to the AAD application
  # -------------------------------------------------------------------------------
  Write-Host "Uploading self-signed certificate to the '$($app.displayName)' application..."
  $addcert = (az ad app credential reset --id $app.appId --cert $keyValue --append | ConvertFrom-Cli)
  Write-Host "Done!"
  return @{"CertThumb" = $thumb; "PfxPassword" = $pwdplain}
}

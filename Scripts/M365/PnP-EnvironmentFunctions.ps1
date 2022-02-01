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
.DESCRIPTION

All environment specific variables and functions are loaded here, to be used by a main function
Adjust all environment specific variables for your DTAP environments.
.EXAMPLE

Load this script file as follows: . .\PnP-EnvironmentFunctions.ps1
#>


function Set-Environment($Environment, $Path, $Site){
  #CSOM first

  # Import Types

  # Import Modules
  If (!(Get-module SharePointPnPPowerShellOnline)) {
    Import-Module SharePointPnPPowerShellOnline -Scope "Local"
  }

	# Set variables
  # NOTE: Stripping comments by the first two replace rows is not needed in Powershell 6+
  $jsonenvironmentFull = Get-Content -Raw -LiteralPath "$($Path)_Environment_$($Environment).jsonc"

  $global:jsonenvironmentMain = (((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  | ConvertFrom-Json).environmentMain

  $global:jsonenvironmentMisc = (((((((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  -replace "{{Name}}","$($global:jsonenvironmentMain.customerName)") `
  -replace "{{Prefix}}","$($global:jsonenvironmentMain.customerPrefix)") `
  -replace "{{O365TenantPrefix}}","$($global:jsonenvironmentMain.customerO365TenantPrefix)") `
  -replace "{{O365GroupsAcceptedEmailDomain}}","$($global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain)") `
  | ConvertFrom-Json).environmentMisc

  $global:jsonsiteSettings = (((((((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  -replace "{{Name}}","$($global:jsonenvironmentMain.customerName)") `
  -replace "{{Prefix}}","$($global:jsonenvironmentMain.customerPrefix)") `
  -replace "{{O365TenantPrefix}}","$($global:jsonenvironmentMain.customerO365TenantPrefix)") `
  -replace "{{O365GroupsAcceptedEmailDomain}}","$($global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain)") `
  | ConvertFrom-Json).siteSettings | Where-Object { $_.name -eq $Site}

  # Set all Service Connection Authentication parameters
  [hashtable]$global:ServiceConnectionMethod = @{}
  foreach ($service in $global:jsonenvironmentMisc.ServiceAuthenticationSchemes.perService) {
    [hashtable]$ServiceConnectionProperties = @{}
    # Set the Authentication Scheme Type Properties
    if ($service.authenticationScheme) {
      $ServiceConnectionProperties += @{"AuthSchemeType" = $service.authenticationScheme}
    }
    else {
      $ServiceConnectionProperties += @{"AuthSchemeType" = $global:jsonenvironmentMisc.ServiceAuthenticationSchemes.default}
    }
    # Set the Authentication Scheme Version Properties
    if ($service.authenticationSchemeVersion) {
      $ServiceConnectionProperties += @{"AuthSchemeVersion" = $service.authenticationSchemeVersion}
    }
    else {
      $ServiceConnectionProperties += @{"AuthSchemeVersion" = ""}
    }
    # Add the full Service Authentication Scheme
    $global:ServiceConnectionMethod += @{$service.serviceName = $ServiceConnectionProperties}
  }

  # Fetch Credentials
  Set-CredentialTargets

  $global:siteUrlTarget = $global:jsonsiteSettings.spurlTarget
  $global:relativeWebTarget = ""
  $global:targetContext = $null
  $global:webTarget = $null
  $global:webTargetUrl = $null
  $global:webTargetRelativeUrl = $null
  $global:siteTargetRelativeUrl = $null
}


function Set-CredentialTargets () {
  $global:dstCred = Get-PnPStoredCredential -Name "$($global:jsonenvironmentMisc.credentialTarget)" -Type PSCredential
  $global:dstGraphCred = Get-PnPStoredCredential -Name "$($global:jsonenvironmentMisc.credentialGraphTarget)" -Type PSCredential
  if ($env:DSTCREDS_USERNAME) {
    #Write-Host "Set Pipeline Cred Properties: $env:DSTCREDS_USERNAME"
    $username = $env:DSTCREDS_USERNAME
    $password = "$env:DSTCREDS_PASSWORD"
    $secstr1 = New-Object -TypeName System.Security.SecureString
    $password.ToCharArray() | ForEach-Object { $secstr1.AppendChar($_) }
    $global:dstCred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $secstr1
  }
  if ($env:DSTCREDS_CLIENTID) {
    #Write-Host "Set Pipeline Cert Properties: $env:DSTCREDS_CLIENTID"
    $clientid = $env:DSTCREDS_CLIENTID
    $thumb = "$env:DSTCREDS_THUMB"
    $secret = "$env:DSTCREDS_SECRET"
    $thumbsecret = $secret
    if ($thumb) {
      $thumbsecret = $thumb + "|" + $thumbsecret
    }
    $secstr2 = New-Object -TypeName System.Security.SecureString
    $thumbsecret.ToCharArray() | ForEach-Object { $secstr2.AppendChar($_) }
    $global:dstGraphCred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $clientid, $secstr2
  }
  [string]$global:dstGraphCredSecret = "$($global:dstGraphCred.GetNetworkCredential().Password)"
  if ($global:dstGraphCred.GetNetworkCredential().Password.Split('|')[1]) {
    [string]$global:dstGraphCredCertThumb = "$($global:dstGraphCred.GetNetworkCredential().Password.Split('|')[0])"
    [string]$global:dstGraphCredSecret = "$($global:dstGraphCred.GetNetworkCredential().Password.Split('|')[1])"
  }
}


function Connect-PnPSpo([object]$properties) {
  if ($global:relativeWebTarget -ne "") {
    $siteurl = "$($global:siteUrlTarget)/$($global:relativeWebTarget)"
  }
  else {
    $siteurl = "$($global:siteUrlTarget)"
  }
  #region DisconnectPnPOnline
  try {
    Write-Host "Disconnecting from current site" -ForegroundColor Green
    Disconnect-PnPOnline -ErrorAction Stop
  }
  catch {
    Write-Host "Not connected" -ForegroundColor "Yellow"
    $Error.Clear()
  }
  #endregion DisconnectPnPOnline
  #region ConnectPnPOnline
  try {
    switch ($properties.AuthSchemeType) {
      "Cred" {
        Write-Host "Connecting to site $siteurl as $($global:dstCred.UserName)" -ForegroundColor "Green"
        Connect-PnPOnline -Url $siteurl -Credentials $global:dstCred -SkipTenantAdminCheck -IgnoreSslErrors -RetryCount 3 -RetryWait 5 -NoTelemetry -WarningAction SilentlyContinue -ErrorAction Stop
      }
      "Cert" {
        if ($global:dstGraphCredCertThumb) {
          Write-Host "Connecting to site $siteurl as $($global:dstGraphCred.UserName)" -ForegroundColor "Green"
          Connect-PnPOnline -Url $siteurl -Tenant $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -ClientId "$($global:dstGraphCred.Username)" -Thumbprint $global:dstGraphCredCertThumb -SkipTenantAdminCheck -IgnoreSslErrors -RetryCount 3 -RetryWait 5 -NoTelemetry -WarningAction SilentlyContinue -ErrorAction Stop
        }
        else {
          Write-Host "No valid 'Certificate Thumbprint' detected. Exiting..."
          exit
        }
      }
      default {
        Write-Host "No Connection type was specified. Exiting..."
        exit
      }
    }
  }
  catch {
    Write-Host "Error connecting to $($siteurl): $($Error[0].ToString())" -ForegroundColor "Red"
    $Error.Clear()
    return
  }
  if (-not (Get-PnPContext)) {
    Write-Host "Error connecting to $($siteurl), unable to establish context" -ForegroundColor "Red"
    $Error.Clear()
    exit
  }
  else {
    try {
      $global:webTarget = Get-PnPWeb -ErrorAction Stop
      $global:targetContext = Get-PnPContext -ErrorAction Stop
      $global:webTargetUrl = $global:webTarget.Url
      $global:webTargetRelativeUrl = $global:webTarget.ServerRelativeUrl
      Write-Host "Connected to Web Target URL: $($global:webTargetUrl)" -ForegroundColor "Green"
      $global:siteTargetRelativeUrl = Get-PnPProperty -ClientObject $global:webTarget.Context.Site -Property ServerRelativeUrl -ErrorAction Stop
      if ($global:siteTargetRelativeUrl -eq "/") {
        $global:siteTargetRelativeUrl = ""
      }
    }
    catch {
      Write-Host "Error connecting to $($siteurl), unable to set additional context sensitive properties: $($Error[0].ToString())" -ForegroundColor "Red"
      $Error.Clear()
      return
    }
  }
  #endregion ConnectPnPOnline
}

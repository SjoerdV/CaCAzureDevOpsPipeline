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

This script will add required app Principals and local certificates for connecting to various API's like the Microsoft Graph and add these Principals to Agent Groups if required
.PARAMETER Environment

Specify the environment this script will run for
.EXAMPLE

Run this script file as follows: .\Install-AzureADAppPrincipals.ps1 -Environment "TEST"
#>
#Get Parameters
[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)] [ValidateSet("TEST", "PROD")][string] $Environment
)
if ($env:ENVIRONMENT) {
  $Environment = $env:ENVIRONMENT
}
if (!$Environment) {
  Write-Host "=== No environment specified. Exiting... ===" -ForegroundColor Red
  return
}
Write-Host "=== This is running for environment: $($Environment) ===" -ForegroundColor Green

#Load Environment Functions
$O365ScriptRootLevel = "..\"
$RepoRootLevel = "$O365ScriptRootLevel..\..\"
. "$($O365ScriptRootLevel)PnP-EnvironmentFunctions.ps1"

## Load Environment
Set-Environment $Environment "$($O365ScriptRootLevel)" "Root"

#Load Helper Functions
. "$($O365ScriptRootLevel)..\Powershell-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)Az-HelperFunctions.ps1"

#Do Stuff
#------------------------------------------------------------------------
# Connect to Az
$global:ServiceConnectionMethod.Az.AuthSchemeType = "Cred"
Connect-Az $global:ServiceConnectionMethod.Az


$appdefinitions = $global:jsonenvironmentMisc.AzureAppsAndPrincipals
foreach ($appdefinition in $appdefinitions) {
  # Add or Update App Principals
  $secret = $null
  $certthumb = $null
  $pfxpwd = $null
  $skipped = $false
  $app = (az ad app list --filter "displayname eq '$($appdefinition.AppName)'" | ConvertFrom-Cli)
  if ($app.Count -gt 1) {
    Write-Host "Multiple Apps with the same name '$($appdefinition.AppName)' detected. This is not supported. Skipping..." -ForegroundColor "Red"
    continue
  }
  [array]$AppAccess = @()
  [array]$AppAccess += $($appdefinition.AppSettings.RequiredResourceAccess)
  $AppAccessJson = (ConvertTo-Json $AppAccess -Depth 10 -Compress) > ".temp-body.json"
  if (!$app.displayName) {
    # Add App Principal
    Write-Host "Creating the Azure AD application and related resources..."
    if ($AppAccess.Count -gt 0) {
      $app = (az ad app create --display-name "$($appdefinition.AppName)" --sign-in-audience "$($appdefinition.AppSettings.SignInAudience)" --required-resource-accesses "@.temp-body.json" | ConvertFrom-Cli)
      $secret = (az ad app credential reset --id $app.appId | ConvertFrom-Cli).password
      if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
        $out = New-AppSelfsignedCertificate $app
        $certthumb = $out.CertThumb
        $pfxpwd = $out.PfxPassword
      }
      $spn = (az ad sp create --id $app.appId | ConvertFrom-Cli)
    }
    else {
      Write-Host "The App Resource Access grants were not loaded. Probably there is an issue with the provided JSON. Exiting...!" -ForegroundColor "Red"
      exit
    }
  }
  else {
    # Update App Principal
    Write-Host "App Principal already exist."
    $Confirm = Read-Host "Do you wish to update the app with id '$($app.appId)' (Y/N)?"
    if($Confirm -match "[y]") {
      Write-Host "Updating..."
      $appupd = (az ad app update --id $app.appId --display-name "$($appdefinition.AppName)" --sign-in-audience "$($appdefinition.AppSettings.SignInAudience)" --required-resource-accesses "@.temp-body.json" | ConvertFrom-Cli)
      $Confirm = $null; $Confirm = Read-Host "You are updating the app with id '$($app.appId)'. Do you wish to generate a new clientsecret (Y/N)?"
      if($Confirm -match "[y]") {
        $secret = (az ad app credential reset --id $app.appId | ConvertFrom-Cli).password
      }
      if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
        $Confirm = Read-Host "You are updating the app with id '$($app.appId)'. Do you wish to generate a new certificate (Y/N)?"
        if($Confirm -match "[y]") {
          $out = New-AppSelfsignedCertificate $app
          $certthumb = $out.CertThumb
          $pfxpwd = $out.PfxPassword
        }
      }
    }
    else {
      Write-Host "Skipping..."
      $skipped = $true
    }
  }
  if (!$skipped) {
    $spn = (az ad sp show --id $app.appId | ConvertFrom-Cli)
    # Start - Add Tags. Known issue: https://github.com/Azure/azure-cli/issues/23027, could still be optimized after July 5 2022.
    try {
      [object]$TagConfig = $appdefinition.AppSettings.ServicePrincipal.TagConfig
      $TagConfigJson = (ConvertTo-Json $TagConfig.tags -Depth 10 -Compress) > ".temp-body-tags.json"
      $updspn = az ad sp update --id $spn.id --set tags="@.temp-body-tags.json" | ConvertFrom-Cli
    }
    catch {
      Write-Host "Failed to update App Tags. Skipping..."
    }
    # End - Add Tags.
    Write-Host "Updating App Settings..."
    $updapp = (az ad app update --id $app.appId --web-redirect-uris @("https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/$($app.appId)/isMSAApp/") | ConvertFrom-Cli)
    Write-Host "Done!"
    # Add SPN to Azure roles
    foreach ($RoleMemberShip in $appdefinition.AppSettings.ServicePrincipal.RoleMemberShips) {
      try {
        Write-Host "Adding App Principal '$($spn.id)' to Azure role '$($RoleMemberShip.DisplayName)'..."
        # elevatate account
        az rest --method post --url "/providers/Microsoft.Authorization/elevateAccess?api-version=2016-07-01"
        # add Azure AD role
        $body = @{
          "roleDefinitionId" = "$($RoleMemberShip.Id)";
          "principalId"      = "$($spn.id)";
          "directoryScopeId" = "/"
        } | ConvertTo-Json -Compress
        $body = $body.Replace('"', '\"')
        $role = (az rest -m post -u "https://graph.microsoft.com/beta/roleManagement/directory/roleAssignments" -b "$body" | ConvertFrom-Cli)
        # remove account elevation
        az role assignment delete --assignee $global:dstCred.UserName --role "User Access Administrator" --scope "/"
        Write-Host "Done!"
      }
      catch {
        if ($Error[0].Exception.Message -notmatch "already exist") {
          Write-Host "Adding App Principal '$($spn.id)' to role '$($RoleMemberShip.DisplayName)' failed: $($Error[0].ToString())" -ForegroundColor "Red"
        }
      }
    }
    # Consent App Permissions
    Write-Host "Consenting App '$($app.appId)' Permissions..."
    $app = (az ad app list --filter "displayname eq '$($appdefinition.AppName)'" | ConvertFrom-Cli)
    $consent = (Start-RetryScriptBlock -ScriptBlock { (az ad app permission admin-consent --id $app.appId | ConvertFrom-Cli) } -Retries 5 -SecondsDelay 5 -Indent " ")
    Write-Host "Done!"
    ## Print results
    Write-Host "================ Secrets ================"
    Write-Host "AppDisplayName              = $($app.displayName)"
    Write-Host "ApplicationId               = $($app.appId)"
    if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
      Write-Host "ApplicationCertThumb        = $(if ($certthumb) { $certthumb } else { "CERT THUMBPRINT ALREADY DOCUMENTED" } )"
      Write-Host "ApplicationPfxPassword      = $(if ($pfxpwd) { $pfxpwd } else { "PFX PASSWORD ALREADY DOCUMENTED" } )"
    }
    Write-Host "ApplicationSecret           = $(if ($secret) { $secret } else { "CLIENT SECRET ALREADY DOCUMENTED" } )"
    Write-Host "TenantName                  = $($global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain)"
    Write-Host "TenantID                    = $($global:jsonenvironmentMisc.AzureADTenantId)"
    Write-Host "================ Secrets ================"
    Write-Host "    SAVE THESE IN A SECURE LOCATION     "
  }
}

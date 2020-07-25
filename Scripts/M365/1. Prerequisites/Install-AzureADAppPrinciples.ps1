<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
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
. "$($O365ScriptRootLevel)Powershell-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)AzureAd-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)PnP-HelperFunctions.ps1"

#Do Stuff
#------------------------------------------------------------------------
# Connect to Azure AD
Connect-Aad "Cred"


# Define Classes
class ResourceAccess {
  [System.String]$Id
  [System.String]$Type
}

class RequiredResourceAccess {
  [System.String]$ResourceAppId
  [System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]]$ResourceAccess
}

class root {
  [System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]]$RequiredResourceAccess
}


# Add App Principals
$SessionInfo = Get-AzureADCurrentSessionInfo

$appdefinitions = $global:jsonenvironmentMisc.AzureAppsAndPrincipals
foreach ($appdefinition in $appdefinitions) {
  $secret = $null
  $certthumb = $null
  $pfxpwd = $null
  $app = Get-AzureADApplication | Where-Object { $_.DisplayName -eq "$($appdefinition.AppName)" }
  $spn = Get-AzureADServicePrincipal -All $true | Where-Object { $_.AppId -eq "$($app.AppId)" }
  $AppAccess = [System.Web.Script.Serialization.JavaScriptSerializer]::new().Deserialize((ConvertTo-Json -Depth 100 -InputObject @($appdefinition.AppSettings.RequiredResourceAccess)), [Microsoft.Open.AzureAD.Model.RequiredResourceAccess[]])
  if (!$app.DisplayName) {
    Write-Host "Creating the Azure AD application and related resources..."
    if ($AppAccess.Count -gt 0) {
      $app = New-AzureADApplication -AvailableToOtherTenants $appdefinition.AppSettings.AvailableToOtherTenants -DisplayName "$($appdefinition.AppName)" -IdentifierUris "https://$($SessionInfo.TenantDomain)/$((New-Guid).ToString())" -RequiredResourceAccess $appAccess -ReplyUrls @("urn:ietf:wg:oauth:2.0:oob")
      $secret = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId -CustomKeyIdentifier "MySecret"
      if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
        $out = New-AppSelfsignedCertificate $app
        $certthumb = $out.CertThumb
        $pfxpwd = $out.PfxPassword
      }
      $spn = New-AzureADServicePrincipal -AppId $app.AppId -DisplayName "$($appdefinition.AppName)" -Tags @($appdefinition.AppSettings.ServicePrincipal.Tags.Name)
      Write-Host "Done!"
    }
    else {
      Write-Host "The App Resource Access grants were not loaded. Probably their is an issue with the provided JSON. Exiting...!" -ForegroundColor "Red"
      exit
    }
  }
  else {
    Write-Host "App Principal already exist. Updating..."
    if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
      $Confirm = Read-Host "You are updating the app with id '$($app.AppId)'. Do you wish to generate a new certificate (Y/N)?"
      if($Confirm -match "[y]") {
        $out = New-AppSelfsignedCertificate $app
        $certthumb = $out.CertThumb
        $pfxpwd = $out.PfxPassword
      }
    }
    Set-AzureADApplication -ObjectId $app.ObjectId -AvailableToOtherTenants $appdefinition.AppSettings.AvailableToOtherTenants -DisplayName "$($appdefinition.AppName)" -RequiredResourceAccess $appAccess -ReplyUrls @("urn:ietf:wg:oauth:2.0:oob")
    Set-AzureADServicePrincipal -ObjectId $spn.ObjectId -AppId $app.AppId -DisplayName "$($appdefinition.AppName)" -Tags @($appdefinition.AppSettings.ServicePrincipal.Tags.Name)
  }
  ## Give consent and print results
  Write-Host "IMPORTANT: Please browse to https://login.microsoftonline.com/$($spn.AppOwnerTenantID)/adminConsent?client_id=$($app.AppId)" -ForegroundColor "Yellow"
  Write-Host "Press any key after auth. An error report about incorrect URIs is expected!"
  [void][System.Console]::ReadKey($true)
  Write-Host "================ Secrets ================"
  Write-Host "LinkedCredential            = $($global:jsonenvironmentMisc.credentialGraphTarget)"
  Write-Host "AppDisplayName              = $($app.DisplayName)"
  Write-Host "ApplicationId               = $($app.AppId)"
  Write-Host "ApplicationSecret           = $(if ($secret.Value) { $secret.Value } else { "CLIENT SECRET ALREADY DOCUMENTED" } )"
  if ($appdefinition.AppSettings.AuthenticationScheme -ne "Secret") {
    Write-Host "ApplicationCertThumb        = $(if ($certthumb) { $certthumb } else { "CERT THUMBPRINT ALREADY DOCUMENTED" } )"
    Write-Host "ApplicationPfxPassword      = $(if ($pfxpwd) { $pfxpwd } else { "PFX PASSWORD ALREADY DOCUMENTED" } )"
  }
  Write-Host "TenantName                  = $($global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain)"
  Write-Host "TenantID                    = $($spn.AppOwnerTenantID)"
  Write-Host "================ Secrets ================"
  Write-Host "    SAVE THESE IN A SECURE LOCATION     "
}

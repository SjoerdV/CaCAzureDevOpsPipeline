<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
.DESCRIPTION

All environment specific variables and functions are loaded here, to be used by a main function
Adjust all environment specific variables for your DTAP environments.
.EXAMPLE

Load this script file as follows: . .\PnP-EnvironmentFunctions.ps1
#>

Set-PnPTraceLog -On -Level Debug

function Set-Environment($Environment, $Path, $Site){
  #CSOM first

  # Import Types

  # Import Modules
  If (!(Get-module SharePointPnPPowerShellOnline)) {
    Import-Module SharePointPnPPowerShellOnline -Scope "Local"
  }

	# Set variables
  # NOTE: Stripping comments by the first two replace rows is not needed in Powershell 6+
  $jsonenvironmentFull = Get-Content -Raw -Path "$($Path)_Environment_$($Environment).jsonc"

  $global:jsonenvironmentMain = (((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  | ConvertFrom-Json).environmentMain

  $global:jsonenvironmentMisc = ((((((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  -replace "{{Name}}","$($global:jsonenvironmentMain.customerName)") `
  -replace "{{Prefix}}","$($global:jsonenvironmentMain.customerPrefix)") `
  -replace "{{O365TenantPrefix}}","$($global:jsonenvironmentMain.customerO365TenantPrefix)") `
  | ConvertFrom-Json).environmentMisc

  $global:jsonsiteSettings = ((((((($jsonenvironmentFull) `
  -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*') `
  -replace '(?ms)/\*.*?\*/') `
  -replace "{{Name}}","$($global:jsonenvironmentMain.customerName)") `
  -replace "{{Prefix}}","$($global:jsonenvironmentMain.customerPrefix)") `
  -replace "{{O365TenantPrefix}}","$($global:jsonenvironmentMain.customerO365TenantPrefix)") `
  | ConvertFrom-Json).siteSettings | Where-Object { $_.name -eq $Site}

  $global:siteUrlTarget = $global:jsonsiteSettings.spurlTarget
  $global:relativeWebTarget = ""
  $global:targetContext = $null
  $global:webTarget = $null
  $global:webTargetUrl = $null
  $global:webTargetRelativeUrl = $null
  $global:siteTargetRelativeUrl = $null
  Set-CredentialTargets
}


function Set-CredentialTargets () {
  $global:dstCred = Get-PnPStoredCredential -Name "$($global:jsonenvironmentMisc.credentialTarget)" -Type PSCredential
  if ($env:DSTCREDS_USERNAME) {
    $username = $env:DSTCREDS_USERNAME
    $password = [System.Environment]::getEnvironmentVariable('DSTCREDS_PASSWORD',[System.EnvironmentVariableTarget]::User)
    $secstr = New-Object -TypeName System.Security.SecureString
    $password.ToCharArray() | ForEach-Object { $secstr.AppendChar($_) }
    $global:dstCred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, $secstr
  }
}


function Connect-SharePointOnline() {
  if ($global:relativeWebTarget -ne "") {
    $siteurl = "$($global:siteUrlTarget)/$($global:relativeWebTarget)"
  }
  else {
    $siteurl = "$($global:siteUrlTarget)"
  }
  #region DisconnectPnPOnline
  try {
    Write-Output "Disconnecting from current site"
    Disconnect-PnPOnline -ErrorAction Stop
  }
  catch {
    Write-Output "Not connected"
    $Error.Clear()
  }
  #endregion DisconnectPnPOnline
  #region ConnectPnPOnline
  try {
    # Refetching credentials to work around an issue with these Azure AD modules interfering with credntials
    # Reference: https://www.benstegink.com/azure-ad-powershell-module-issue/#.XQoD7W5uK9w
    Set-CredentialTargets
    Write-Output "Connecting to site $siteurl as $($global:dstCred.UserName)"
    Connect-PnPOnline -Url $siteurl -Credentials $global:dstCred -RetryCount 3 -RetryWait 5 -NoTelemetry -ErrorAction Stop
    if (-not (Get-PnPContext)) {
      Write-Output "Error connecting to $($siteurl), unable to establish context"
      return
    }
    else {
      $global:webTarget = Get-PnPWeb -ErrorAction Stop
      $global:targetContext = Get-PnPContext -ErrorAction Stop
      $global:webTargetUrl = $global:webTarget.Url
      $global:webTargetRelativeUrl = $global:webTarget.ServerRelativeUrl
      Write-Output "Connected to Web Target URL: $($global:webTargetUrl)"
      $global:siteTargetRelativeUrl = Get-PnPProperty -ClientObject $global:webTarget.Context.Site -Property ServerRelativeUrl -ErrorAction Stop
      if ($global:siteTargetRelativeUrl -eq "/") {
        $global:siteTargetRelativeUrl = ""
      }
    }
  }
  catch {
    Write-Output "Error connecting to $($siteurl): $($Error[0].ToString())"
    $Error.Clear()
    return
  }
  #endregion ConnectPnPOnline
}

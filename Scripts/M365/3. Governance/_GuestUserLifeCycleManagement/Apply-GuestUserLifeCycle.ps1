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

This script needs to be scheduled daily, and will manage each Guest User lifecycle: Creation, Expiration, Re-activation and Deletion.
.PARAMETER Environment

Specify the environment this script will run for
.EXAMPLE

Run this script file as follows: .\Apply-GuestUserLifeCycle.ps1 -Environment "TEST"
#>
#Get Parameters
[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)] [ValidateSet("TEST", "PROD")][string] $Environment
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
$O365ScriptRootLevel = "..\..\"
$RepoRootLevel = "$O365ScriptRootLevel..\..\"
. "$($O365ScriptRootLevel)PnP-EnvironmentFunctions.ps1"

## Load Environment
Set-Environment $Environment "$($O365ScriptRootLevel)" "Root"

#Load Helper Functions
. "$($O365ScriptRootLevel)..\Powershell-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)PnP-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)Mg-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)Exo-HelperFunctions.ps1"

#Do Stuff
#------------------------------------------------------------------------
# Connect to Microsoft Graph
Connect-Mg $global:ServiceConnectionMethod.Mg


# Connect to Exchange Online
Connect-Exo $global:ServiceConnectionMethod.Exo


# Set Variables Processing
$Today = (Get-Date)
$StaleAgeInDays = 180 # when should an External Account expire
$DeleteAgeInDays = 360 # when should an External Account be deleted, when not Reactivated. should be greater than StaleAgeInDays


# Set Variables SharePoint
$provisioningUrl = "$($global:jsonenvironmentMisc.tenantUrl)"
$ProvisioningListName = "Expired Guest User List"


# Fetch any recent user additions
$EndDate = (Get-Date).AddDays(1)
$StartDate = (Get-Date).AddDays(-90)
$Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations "Add User" -ResultSize 2000 -Formatted)
$ToBeProcessed = @()
ForEach ($Rec in $Records) {
  $AuditData = ConvertFrom-Json $Rec.Auditdata
  # Only process the additions of guest users to groups
  If ($AuditData.ObjectId -Like "*#EXT#*") {
    # Do not add duplicate guests
    if ($ToBeProcessed.Guest -notcontains $AuditData.ObjectId) {
      $ToBeProcessed += @{Guest = $AuditData.ObjectId; Actor = $AuditData.UserId; Created = (Get-Date $AuditData.CreationTime)}
    }
  }
}


# Add Owner to CustomAttribute14
ForEach ($Object in $ToBeProcessed) {
  if (!(Get-MailUser $Object.Guest).CustomAttribute14) {
    Write-Host "Setting '$($Object.Actor)' as owner for guest '$($Object.Guest)'"
    Set-MailUser -Identity $Object.Guest -CustomAttribute14 $Object.Actor
  }
}


# Expire accounts or set expiration date in CustomAttribute15
# Get endpoint
$uri = "/users/?`$select=accountEnabled,userPrincipalName,userType&`$filter=accountEnabled eq true and userType eq `'Guest`'"
# place the call
[array]$GuestUsersMG = @(Start-RetryScriptBlock -ScriptBlock { Invoke-EasyGraphRequest -Resource "$($uri)" -Method "GET" -APIVersion "beta" -ContentType "application/json" -ErrorAction Stop } -Retries 5 -SecondsDelay 5)
foreach ($GuestUser in $GuestUsersMG) {
  $MailUser = Get-MailUser -Identity $GuestUser.userPrincipalName
  $Owner = $MailUser.CustomAttribute14
  $ExpirationDate = $MailUser.CustomAttribute15
  if ($ExpirationDate) {
    if ($Today -ge (Get-Date $ExpirationDate)) {
      #Expire existing accounts with current date is greater than expiration date
      try {
        Write-Host "Disabling guest '$($GuestUser.userPrincipalName)'"
        ## construct request body
        $body = @{
          accountEnabled = $false
        }
          # Get endpoint
        $uri = "/users/$([System.Web.HttpUtility]::UrlEncode($GuestUser.userPrincipalName))"
        # place the call
        $result = Start-RetryScriptBlock -ScriptBlock { Invoke-EasyGraphRequest -Resource "$($uri)" -Method "PATCH" -APIVersion "beta" -Body $body -ContentType "application/json" -ErrorAction Stop } -Retries 5 -SecondsDelay 5
        Write-Host "Done!"
      }
      catch {
        Write-Host "Disabling guest '$($GuestUser.userPrincipalName)' failed: $($Error[0].ToString())" -ForegroundColor "Red"
      }
      #Write disabled account to SharePoint List
      Write-Host "Add expired guest '$($GuestUser.userPrincipalName)' to SharePoint List"
      Add-GuestExpirationToSharePointList $GuestUser.userPrincipalName #This function should add a SharePoint List Item where the Title column equals '$GuestUser.userPrincipalName' to the SharePoint List
      #Send email to owner and admin
      $Subject = "External User $($GuestUser.userPrincipalName) has expired"
      $Body = "The account for External User $($GuestUser.userPrincipalName) with Owner: '$($Owner)' has expired, by mandatory automatic expiration process. If you wish to reactivate this account go to the 'Expired Guest User List' and click the Reactivate button."
      Write-Host "Sending email to owner $($Owner)"
      Send-PnPMail -To $Owner -Subject "$($Subject)" -Body "$($Body)"
      Write-Host "Sending email to admin $($global:dstCred.Username)"
      Send-PnPMail -To $global:dstCred.Username -Subject "$($Subject)" -Body "$($Body)"
      #Set a script scoped variable with the last processed guest user
      $script:LastGuestUpnExpired = "$($GuestUser.userPrincipalName)"
    }
  }
  else {
    $ExpirationDate = $Today.AddDays($StaleAgeInDays).toString('u')
    #Add expiration date (currentdate + stale age) to new accounts
    Write-Host "Setting '$ExpirationDate' as expiration date for guest '$($GuestUser.userPrincipalName)'"
    Set-MailUser -Identity $GuestUser.userPrincipalName -CustomAttribute15 $ExpirationDate
  }
}

# Reactivate Guest User Accounts
Write-Host "Fetching Stale Guest Users..."
$GuestUsersSPO = Get-GuestReactivationsFromSharePointList #This function should return a collection of SharePoint List Items with all necessary UPN values contained in the Title column.
foreach ($GuestUser in $GuestUsersSPO) {
  $MailUser = Get-MailUser -Identity $GuestUser
  $Owner = $MailUser.CustomAttribute14
  $ExpirationDate = $MailUser.CustomAttribute15
  #Reactivate previously expired account
  try {
    Write-Host "Reactivating Stale Guest User '$($GuestUser.Title)'..."
    ## construct request body
    $body = @{
      accountEnabled = $true
    }
    # Get endpoint
    $uri = "/users/$([System.Web.HttpUtility]::UrlEncode($GuestUser.Title))"
    # place the call
    $result = Start-RetryScriptBlock -ScriptBlock { Invoke-EasyGraphRequest -Resource "$($uri)" -Method "PATCH" -APIVersion "beta" -Body $body -ContentType "application/json" -ErrorAction Stop } -Retries 5 -SecondsDelay 5
    Write-Host "Done!"
  }
  catch {
    Write-Host "Disabling Stale Guest User '$($GuestUser.Title)' failed: $($Error[0].ToString())" -ForegroundColor "Red"
  }
  #Set new expiration date
  try {
    $ExpirationDate = $Today.AddDays($StaleAgeInDays).toString('u')
    Write-Host "Setting Expiry Date '$ExpirationDate' on Guest User '$($GuestUser.Title)'..."
    Set-MailUser -Identity $GuestUser.Title -CustomAttribute15 $ExpirationDate
    Write-Host "Done!"
  }
  catch {
    Write-Host "Setting Reactivated Guest User '$($GuestUser.Title)' New Expiry Date '$ExpirationDate' failed: $($Error[0].ToString())" -ForegroundColor "Red"
  }
  #Remove reactivated account from SharePoint List
  Write-Host "Remove entry for reactivated account for Guest User '$($GuestUser.Title)' from SharePoint List..."
  Remove-GuestFromSharePointList $GuestUser.Title #This function should remove the SharePoint List Item where the Title column equals '$GuestUser.Title' from the SharePoint List
  #Send email to owner and admin
  $Subject = "External User $($GuestUser.Title) is reactivated"
  $Body = "The account for External User '$($GuestUser.Title)' with Owner: '$($Owner)' has been reactivated by a user initiated reactivatiion process. The new Expiration date is '$ExpirationDate'"
  Write-Host "Sending email to owner $($Owner)"
  Send-PnPMail -To $Owner -Subject "$($Subject)" -Body "$($Body)"
  Write-Host "Sending email to admin $($global:dstCred.Username)"
  Send-PnPMail -To $global:dstCred.Username -Subject "$($Subject)" -Body "$($Body)"
  ## Set a script scoped variable with the last processed guest user
  $script:LastGuestUpnReactivated = $GuestUser
}

# Delete Guest User Accounts
[int]$StaleDeleteDifference = $DeleteAgeInDays - $StaleAgeInDays
# Get endpoint
$uri = "/users/?`$select=accountEnabled,userPrincipalName,userType&`$filter=accountEnabled eq false and userType eq `'Guest`'"
# place the call
[array]$GuestUsersMG = @(Start-RetryScriptBlock -ScriptBlock { Invoke-EasyGraphRequest -Resource "$($uri)" -Method "GET" -APIVersion "beta" -ContentType "application/json" -ErrorAction Stop } -Retries 5 -SecondsDelay 5)
foreach ($GuestUser in $GuestUsersMG) {
  $MailUser = Get-MailUser -Identity $GuestUser.userPrincipalName
  $Owner = $MailUser.CustomAttribute14
  $ExpirationDate = $MailUser.CustomAttribute15
  if ($Today -ge (Get-Date $ExpirationDate).addDays($StaleDeleteDifference)) {
    #Remove account when current date is greater than expiration date + (deletion age - stale age)
    Remove-MailUser -Identity $GuestUser.userPrincipalName -Confirm:$false
    #Remove SharePoint List item, if exists
    Write-Host "Remove entry for deleted account for Guest User '$($GuestUser.Title)' from SharePoint List..."
    Remove-GuestFromSharePointList $GuestUser.userPrincipalName
    #Send email to owner and admin
    $Subject = "External User $($GuestUser.userPrincipalName) is permanently deleted"
    $Body = "The account for External User '$($GuestUser.userPrincipalName)' with Owner: '$($Owner)' has been permanently deleted."
    Write-Host "Sending email to owner $($Owner)"
    Send-PnPMail -To $Owner -Subject "$($Subject)" -Body "$($Body)"
    Write-Host "Sending email to admin $($global:dstCred.Username)"
    Send-PnPMail -To $global:dstCred.Username -Subject "$($Subject)" -Body "$($Body)"
  }
}

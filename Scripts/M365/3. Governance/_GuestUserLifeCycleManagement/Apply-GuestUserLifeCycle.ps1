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
$GuestUsersEXO = Get-User -RecipientTypeDetails GuestMailUser -ResultSize Unlimited | Where-Object { !$_.AccountDisabled }
foreach ($GuestUser in $GuestUsersEXO) {
  $MailUser = Get-MailUser -Identity $GuestUser.UserPrincipalName
  $Owner = $MailUser.CustomAttribute14
  $ExpirationDate = $MailUser.CustomAttribute15
  if ($ExpirationDate) {
    if ($Today -ge (Get-Date $ExpirationDate)) {
      #Expire existing accounts with current date is greater than expiration date
      try {
        Write-Host "Disabling guest '$($GuestUser.UserPrincipalName)'"
        ## construct request body
        $body = @{
          accountEnabled = $false
        }
          # Get endpoint
        $uri = "/users/$([System.Web.HttpUtility]::UrlEncode($GuestUser.UserPrincipalName))"
        # place the call
        $result = Start-RetryScriptBlock -ScriptBlock { Invoke-EasyGraphRequest -Resource "$($uri)" -Method "PATCH" -APIVersion "beta" -Body $body -ContentType "application/json" -ErrorAction Stop } -Retries 5 -SecondsDelay 5
        Write-Host "Done!"
      }
      catch {
        Write-Host "Disabling guest '$($GuestUser.UserPrincipalName)' failed: $($Error[0].ToString())" -ForegroundColor "Red"
      }
      #Write disabled account to SharePoint List
      Write-Host "Add expired guest '$($GuestUser.UserPrincipalName)' to SharePoint List"
      Add-GuestExpirationToSharePointList $GuestUser.UserPrincipalName
      #Send email to owner and admin
      $Subject = "External User $($GuestUser.UserPrincipalName) has expired"
      $Body = "The account for External User $($GuestUser.UserPrincipalName) with Owner: '$($Owner)' has expired, by mandatory automatic expiration process. If you wish to reactivate this account go to the 'Expired Guest User List' and click the Reactivate button."
      Write-Host "Sending email to owner $($Owner)"
      Send-Report $Owner $Subject $Body
      Write-Host "Sending email to admin $($global:dstCred.Username)"
      Send-Report $global:dstCred.Username $Subject $Body
      #Set a script scoped variable with the last processed guest user
      $script:LastGuestUpnExpired = "$($GuestUser.UserPrincipalName)"
    }
  }
  else {
    $ExpirationDate = $Today.AddDays($StaleAgeInDays).toString('u')
    #Add expiration date (currentdate + stale age) to new accounts
    Write-Host "Setting '$ExpirationDate' as expiration date for guest '$($GuestUser.UserPrincipalName)'"
    Set-MailUser -Identity $GuestUser.UserPrincipalName -CustomAttribute15 $ExpirationDate
  }
}

#Wait until expirations are synced to Exchange Online
if ($script:LastGuestUpnExpired) {
  Start-WaitOnMailUserAccountDisabledStatus $script:LastGuestUpnExpired "GuestMailUser" $true
}

# Reactivate Guest User Accounts
Write-Host "Fetching Stale Guest Users..."
$GuestUsers = Get-GuestReactivationsFromSharePointList #This function should return a collection of SharePoint List Items with all necessary UPN values contained in the Title column.
foreach ($GuestUser in $GuestUsers) {
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
  Remove-GuestFromSharePointList $GuestUser.Title
  #Send email to owner and admin
  $Subject = "External User $($GuestUser.Title) is reactivated"
  $Body = "The account for External User '$($GuestUser.Title)' with Owner: '$($Owner)' has been reactivated by a user initiated reactivatiion process. The new Expiration date is '$ExpirationDate'"
  Write-Host "Sending email to owner $($Owner)"
  Send-Report $Owner $Subject $Body
  Write-Host "Sending email to admin $($global:dstCred.Username)"
  Send-Report $global:dstCred.Username $Subject $Body
  ## Set a script scoped variable with the last processed guest user
  $script:LastGuestUpnReactivated = $GuestUser
}

#Wait until expirations are synced to Exchange Online
if ($script:LastGuestUpnReactivated) {
  Start-WaitOnMailUserAccountDisabledStatus $script:LastGuestUpnReactivated "GuestMailUser" $false
}

# Delete Guest User Accounts
[int]$StaleDeleteDifference = $DeleteAgeInDays - $StaleAgeInDays
$GuestUsersEXO = Get-User -RecipientTypeDetails GuestMailUser -ResultSize Unlimited | Where-Object { $_.AccountDisabled }
foreach ($GuestUser in $GuestUsersEXO) {
  if ($Today -ge (Get-Date ((Get-MailUser $GuestUser.UserPrincipalName).CustomAttribute14)).addDays($StaleDeleteDifference)) {
    $MailUser = Get-MailUser -Identity $GuestUser.UserPrincipalName
    $Owner = $MailUser.CustomAttribute14
    $ExpirationDate = $MailUser.CustomAttribute15
    #Remove account when current date is greater than expiration date + (deletion age - stale age)
    Remove-MailUser -Identity $GuestUser.UserPrincipalName -Confirm:$false
    #Remove SharePoint List item, if exists
    Write-Host "Remove entry for deleted account for Guest User '$($GuestUser.Title)' from SharePoint List..."
    Remove-GuestFromSharePointList $GuestUser.UserPrincipalName
    #Send email to owner and admin
    $Subject = "External User $($GuestUser.UserPrincipalName) is permanently deleted"
    $Body = "The account for External User '$($GuestUser.UserPrincipalName)' with Owner: '$($Owner)' has been permanently deleted."
    Write-Host "Sending email to owner $($Owner)"
    Send-Report $Owner $Subject $Body
    Write-Host "Sending email to admin $($global:dstCred.Username)"
    Send-Report $global:dstCred.Username $Subject $Body
  }
}

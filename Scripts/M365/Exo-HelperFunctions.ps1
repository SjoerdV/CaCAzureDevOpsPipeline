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

Load this script file as follows: . .\Exo-HelperFunctions.ps1
#>


# Import Modules
If ($global:ServiceConnectionMethod.Exo.AuthSchemeVersion -eq "V2") {
  If (!(Get-Module ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement -DisableNameChecking -Force
  }
}


# Add Functions
function Connect-Exo([object]$properties) {
  if ($properties.AuthSchemeVersion -eq "V1") {
    # make sure the connected user is Powershell enabled: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/disable-access-to-exchange-online-powershell?view=exchange-ps
    try {
      # Clear existing Sessions in variable
      if ($Session) {
        Remove-PSSession $Session -ErrorAction Ignore
        $Error.Clear()
      }
      # Clear WinRM Sessions
      $sessions = Get-PSSession -ErrorAction Ignore | Where-Object { $_.ComputerName -match "outlook.office365.com" }
      $Error.Clear()
      foreach ($s in $sessions) {
        Remove-PSSession $s -ErrorAction Ignore
        $Error.Clear()
      }
      # Connect to Exchange Online Connector
      $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:dstCred -Authentication Basic -AllowRedirection
      Import-PSSession $Session -DisableNameChecking -AllowClobber
      # Prevent Exchange Online Credential Prompt
      # Reference: https://www.lieben.nu/liebensraum/2018/01/exchange-online-reconnect-script-v2/
      Write-Host "Connected to Exchange Online, exporting module..."
      $temporaryModulePath = (Join-Path $Env:TEMP -ChildPath "temporaryEXOModule")
      $res = Export-PSSession -Session $Session -CommandName * -OutputModule $temporaryModulePath -AllowClobber -Force
      $temporaryModulePath = Join-Path $temporaryModulePath -ChildPath "temporaryEXOModule.psm1"
      Write-Host "Rewriting Exchange Online module, please wait..."
      $regex = '^.*\bhost\.UI\.PromptForCredential\b.*$'
      (Get-Content $temporaryModulePath) -replace $regex, "-Credential `$global:dstCred ``" | Set-Content $temporaryModulePath
      $Session | Remove-PSSession -Confirm:$False
      Write-Host "Module rewritten, re-importing..."
      Import-Module -Name $temporaryModulePath -DisableNameChecking -WarningAction SilentlyContinue -Force
      Write-Host "Module imported, you may now use all Exchange Online commands"
      return $temporaryModulePath
    }
    catch {
      Write-Host "Error Connecting to Exchange Online Connector: $($Error[0].ToString())" -ForegroundColor "Red"
      exit
    }
  }
  if ($properties.AuthSchemeVersion -eq "V2") {
    try {
      switch ($properties.AuthSchemeType) {
        "Cred" {
          Write-Host "Connecting to Exchange Online as $($global:dstCred.UserName)" -ForegroundColor "Green"
          $exo = Connect-ExchangeOnline -ShowBanner:$false -Credential $global:dstCred -Organization $global:jsonenvironmentMisc.AzureADTenantId -ErrorAction Stop
        }
        "Thumb" {
          if ($global:dstGraphCredCertThumb) {
            Write-Host "Connecting to Exchange Online as $($global:dstGraphCred.UserName)" -ForegroundColor "Green"
            $exo = Connect-ExchangeOnline -ShowBanner:$false -AppId "$($global:dstGraphCred.Username)" -CertificateThumbprint $global:dstGraphCredCertThumb -Organization $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -ErrorAction Stop
          }
          else {
            Write-Host "No valid 'Pfx Password' detected. Exiting..."
            exit
          }
        }
        "PfxFile" {
          if ($global:dstGraphCredCertPfxPassword) {
            Write-Host "Connecting to Exchange Online as $($global:dstGraphCred.UserName) with Certificate Pfx File" -ForegroundColor "Green"
            $exo = Connect-ExchangeOnline -ShowBanner:$false -AppId "$($global:dstGraphCred.Username)" -CertificateFilePath "$($global:dstGraphCredCertPfxFilePath)" -CertificatePassword $global:dstGraphCredCertPfxPassword -Organization $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -ErrorAction Stop
          }
          else {
            Write-Host "No valid 'Certificate Pfx File' detected. Exiting..."
            exit
          }
        }
      }
    }
    catch {
      Write-Host "Connecting to Exchange Online failed: $($Error[0].ToString())" -ForegroundColor "Red"
      exit
    }
  }
}


function Start-WaitOnMailUserAccountDisabledStatus([string]$Upn, [string]$Type, [bool]$Status) {
  ## Wait for Mail User '$($Upn)' with RecipientTypeDetails equals '$Type' AccountDisabled equals '$Status'
  try {
    $MailUser = $null
    $Index = 1
    $Max = 30
    while (!$MailUser -and $Index -le $Max) {
      $MailUser = Get-User -RecipientTypeDetails $Type -ResultSize Unlimited | Where-Object { $_.AccountDisabled -eq $Status -and $_.UserPrincipalName -eq "$($Upn)" } -ErrorAction Ignore
      Write-Host "  Waiting for Mail User '$($Upn)' with RecipientTypeDetails equals '$Type' AccountDisabled equals '$Status'..."
      if ($Index -eq $Max - 10) {
        Write-Host "  Trying reconnect to Exchange Online..."
        Connect-Exo $global:ServiceConnectionMethod.Exo
      }
      Start-Sleep -Seconds 10
      $Index++
    }
    if (!$MailUser -or $Index -ge $Max) {
      throw "Mail User '$($Upn)' with RecipientTypeDetails equals '$Type' AccountDisabled equals '$Status' was not found. Terminating..."
    }
  }
  catch {
    Write-Host "Waiting on Mail User '$($Upn)' with RecipientTypeDetails equals '$Type' AccountDisabled equals '$Status' failed: $($Error[0].ToString())" -ForegroundColor Red
    exit
  }
}

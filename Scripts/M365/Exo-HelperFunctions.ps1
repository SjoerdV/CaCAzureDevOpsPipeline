<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
All helper functions for the Exchange Online module are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\Exo-HelperFunctions.ps1
#>


# Import Modules
if ($(($global:jsonenvironmentMisc.ServiceAuthenticationSchemes.perService | Where-Object {$_.serviceName -eq 'Exo'}).authenticationSchemeVersion) -eq "Cert") {
  If (!(Get-module ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement -DisableNameChecking -Force
  }
}


# Add Functions
function Connect-Exo([string]$type,[string]$version) {
  if ($version -eq "V1") {
    # make sure the connected user is Powershell enabled: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/disable-access-to-exchange-online-powershell?view=exchange-ps
    try {
      # Clear existing Sessions in variable
      if ($Session) {
        Remove-PSSession $Session -ErrorAction Ignore
        $Error.Clear()
      }
      # Clear WinRM Sessions
      $sessions = Get-PSSession -ErrorAction Ignore | Where-Object {$_.ComputerName -match "outlook.office365.com"}
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
      write-logfile "Connected to Exchange Online, exporting module..."
      $temporaryModulePath = (Join-Path $Env:TEMP -ChildPath "temporaryEXOModule")
      $res = Export-PSSession -Session $Session -CommandName * -OutputModule $temporaryModulePath -AllowClobber -Force
      $temporaryModulePath = Join-Path $temporaryModulePath -ChildPath "temporaryEXOModule.psm1"
      write-logfile "Rewriting Exchange Online module, please wait..."
      $regex='^.*\bhost\.UI\.PromptForCredential\b.*$'
      (Get-Content $temporaryModulePath) -replace $regex, "-Credential `$global:dstCred ``" | Set-Content $temporaryModulePath
      $Session | Remove-PSSession -Confirm:$False
      write-logfile "Module rewritten, re-importing..."
      Import-Module -Name $temporaryModulePath -DisableNameChecking -WarningAction SilentlyContinue -Force
      write-logfile "Module imported, you may now use all Exchange Online commands"
      return $temporaryModulePath
    }
    catch {
      write-logfile "Error Connecting to Exchange Online Connector: $($Error[0].ToString())" "Red" $true
      exit
    }
  }
  if ($version -eq "V2") {
    Import-Module ExchangeOnlineManagement
    try {
      switch ($type) {
        "Cred" {
          write-logfile "Connecting to Exchange Online as $($global:dstCred.UserName)" "Green"
          $exo = Connect-ExchangeOnline -ShowBanner:$false -Credential $global:dstCred -Organization $global:jsonenvironmentMisc.AzureADTenantId -ErrorAction Stop
        }
        "Cert" {
          if ($global:dstGraphCredCertThumb) {
            write-logfile "Connecting to Exchange Online as $($global:dstGraphCred.UserName)" "Green"
            $certpath = (Get-ChildItem -Path "$(Resolve-Path -Path $($O365ScriptRootLevel))" -Include "$($global:jsonenvironmentMain.customerPrefix.ToLower())-$($global:dstGraphCred.UserName)-$Environment.pfx" -Recurse | Select-Object -First 1).FullName
            $exo = Connect-ExchangeOnline -ShowBanner:$false -AppId "$($global:dstGraphCred.Username)" -CertificateThumbPrint $global:dstGraphCredCertThumb -Organization $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -ErrorAction Stop
          }
          else {
            write-logfile "No valid 'Pfx Password' detected. Exiting..."
            exit
          }
        }
      }
    }
    catch {
      write-logfile "Connecting to Exchange Online failed: $($Error[0].ToString())" "Red" $true
      exit
    }
  }
}

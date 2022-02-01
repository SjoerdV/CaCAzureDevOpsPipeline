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
All helper functions for the Azure Active Directory module are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\ActiveDirectory-HelperFunctions.ps1
#>

# Import Modules
If (!(Get-module AzureADPreview)) {
  Import-Module AzureADPreview -DisableNameChecking -Force
}


function Connect-Aad([object]$properties) {
  try {
    switch ($properties.AuthSchemeType) {
      "Cred" {
        Write-Host "Connecting to Azure Active Directory as $($global:dstcred.UserName)" -ForegroundColor "Green"
        $aad = Connect-AzureAD -Credential $global:dstCred -TenantId $global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain -ErrorAction Stop
      }
      "Cert" {
        if ($global:dstGraphCredCertThumb) {
          Write-Host "Connecting to Azure Active Directory as $($global:dstGraphCred.UserName)" -ForegroundColor "Green"
          $add = Connect-AzureAD -TenantId "$($global:jsonenvironmentMain.customerO365GroupsAcceptedEmailDomain)" -ApplicationId "$($global:dstGraphCred.Username)" -CertificateThumbprint $global:dstGraphCredCertThumb
        }
        else {
          Write-Host "No valid 'Certificate Thumbprint' detected. Exiting..." -ForegroundColor "Red"
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
    Write-Host "Connecting to Azure AD failed: $($Error[0].ToString())" -ForegroundColor "Red"
    exit
  }
}

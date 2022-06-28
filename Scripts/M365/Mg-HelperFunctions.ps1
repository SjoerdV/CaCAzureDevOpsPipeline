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

Load this script file as follows: . /Mg-HelperFunctions.ps1
#>

# Import Modules
If (!(Get-Module EasyGraph)) {
  Import-Module EasyGraph -DisableNameChecking -Force
}

# Select Graph Profile
#Select-MgProfile -Name "v1.0"

function Connect-Mg([object]$properties) {
  try {
    switch ($properties.AuthSchemeType) {
      "Cred" {
        Write-Host "Connecting to Microsoft Graph with a user account '$($global:dstcred.UserName)' is not supported" -ForegroundColor "Green"
        exit
      }
      "Thumb" {
        if ($global:dstGraphCredCertThumb) {
          Write-Host "Connecting to Microsoft Graph as $($global:dstGraphCred.UserName) with Certificate Thumbprint" -ForegroundColor "Green"
          $mg = Connect-EasyGraph -TenantId "$($global:jsonenvironmentMisc.AzureADTenantId)" -AppId "$($global:dstGraphCred.Username)" -Thumbprint $global:dstGraphCredCertThumb -ErrorAction Stop
        }
        else {
          Write-Host "No valid 'Certificate Thumbprint' detected. Exiting..." -ForegroundColor "Red"
          exit
        }
      }
      "PfxFile" {
        if ($global:dstGraphCredCertPfxPassword) {
          Write-Host "Connecting to Microsoft Graph as $($global:dstGraphCred.UserName) with Certificate Pfx File" -ForegroundColor "Green"
          $mg = Connect-EasyGraph -TenantId "$($global:jsonenvironmentMisc.AzureADTenantId)" -AppId "$($global:dstGraphCred.Username)" -PfxFilePath "$($global:dstGraphCredCertPfxFilePath)" -PfxPassword $global:dstGraphCredCertPfxPassword -ErrorAction Stop
        }
        else {
          Write-Host "No valid 'Certificate Pfx File' detected. Exiting..."
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
    Write-Host "Connecting to Microsoft Graph failed: $($Error[0].ToString())" -ForegroundColor "Red"
    exit
  }
}

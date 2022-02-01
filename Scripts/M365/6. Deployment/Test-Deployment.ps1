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

This script will just make a connection to SharePoint and is used to Test a pipeline deployment.
.PARAMETER Environment

Specify the environment this script will run for
.EXAMPLE

Run this script file as follows: .\Test-Deployment.ps1 -Environment "TEST"
#>
#Example:
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
$O365ScriptRootLevel = "..\"
$RepoRootLevel = "$O365ScriptRootLevel..\..\"
. "$($O365ScriptRootLevel)PnP-EnvironmentFunctions.ps1"

## Load Environment
Set-Environment $Environment "$($O365ScriptRootLevel)" "Root"

#Load Helper Functions
. "$($O365ScriptRootLevel)Powershell-HelperFunctions.ps1"
. "$($O365ScriptRootLevel)PnP-HelperFunctions.ps1"

#Do Stuff
#------------------------------------------------------------------------
# Configure Site Collection

## Connect to SharePoint Online
Connect-PnPSpo $global:ServiceConnectionMethod.PnPSpo

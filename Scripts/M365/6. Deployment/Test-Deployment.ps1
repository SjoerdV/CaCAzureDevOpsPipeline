<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
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

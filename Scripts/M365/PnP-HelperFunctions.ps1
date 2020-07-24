<#
.SYNOPSIS

By: Sjoerd de Valk, SP de Valk Consultancy, 2020
All helper functions for the SharePoint PnP module are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\PnP-HelperFunctions.ps1
#>

# Import Modules
If (!(Get-module SharePointPnPPowerShellOnline)) {
  Import-Module SharePointPnPPowerShellOnline -Scope "Local" -DisableNameChecking -ErrorAction SilentlyContinue
}

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
All helper functions for the SharePoint PnP module are loaded here, to be used by a main function
.DESCRIPTION

Do not adjust any function. Just add new ones. Clean-up later.
.EXAMPLE

Load this script file as follows: . .\PnP-HelperFunctions.ps1
#>

# Import Functions
function Add-GuestExpirationToSharePointList([string]$GuestUpn) {
  Write-Host "Not implemented"
}


function Get-GuestReactivationsFromSharePointList() {
  Write-Host "Not implemented"
}


function Remove-GuestFromSharePointList([string]$GuestUpn) {
  Write-Host "Not implemented"
}

---
title:  'Guest User LifeCycle Management'
author:
- Sjoerd de Valk, SPdeValk Consultancy
date: 2020-07-25T16:00:00+02:00
last_modified_at: 2020-07-25T16:00:00+02:00
keywords: [azure, devops, pipeline, yaml, microsoft365]
abstract: |
  This document is about managing Microsoft 365 Azure B2B Guest Users using a Azure DevOps YAML pipeline.
permalink: /gulm.html
---
## Guest User LifeCycle Management

### Summary

This document is about managing Microsoft 365 Azure B2B Guest Users using a Azure DevOps YAML pipeline.

This document will give manual instructions on adding prerequisite assets to your Microsoft Cloud environment which you can convert to scripted instructions and place in your forked (or another of your choosing) Configuration-as-Code repository.

### Known Issues

None

### Requirements

* Your own Azure DevOps organization, preferably linked to your Azure AD organization.
* If the PowerShell script `Scripts\M365\3. Governance\_GuestUserLifeCycleManagement\Apply-GuestUserLifeCycle.ps1` need to be run locally:
  * Only Windows 7+ with WMI 5.1 and .NET Framework 4.6.1+ is supported (Prefer Windows 10)
  * [PnP PowerShell](https://github.com/pnp/PnP-PowerShell#installation) module needs to be installed
  * [Azure AD Preview](https://www.powershellgallery.com/packages/AzureADPreview/2.0.2.105) module needs to be installed
  * [Exchange Online PowerShell V2](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-v2-module) module needs to be installed  
* If you intend to use the 'Cert' method for authenticating:
  * Make sure you have [previously executed](README.md#add-certificates-and-credentials) the procedure for using the repository and pipeline in your own solution

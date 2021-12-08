# Boehringer.ITEDS.SharePoint.PowerShell

The *Boehringer.ITEDS.SharePoint.PowerShell* is a collection of PowerShell functions which will be used quiet often in all scripts developed for IT EDS SharePoint Online.

## Overview

![Version]
![Language]
[![Platform]][Bitbucket Repo]

To get access to the *Resource Group* in Azure you have to do the request in MyService [Privileged or additional Access - Grant](https://boehringer.service-now.com/bi?id=sc_cat_item&sys_id=0e8963abdb18cc50497151c6f496195e) for the **Configuration Item** `BI-CS-O365.SHAREPOINT-ONLINE(PRODUCTIVE)` and the granted group `RG-M365-Automation`. This request must be approved by the *System Lead* of *BI-CS-O365.SHAREPOINT-ONLINE*.

## Table of Contents

- [Boehringer.ITEDS.SharePoint.PowerShell](#boehringeritedssharepointpowershell)
  - [Overview](#overview)
  - [Table of Contents](#table-of-contents)
  - [Modules](#modules)
  - [Installation](#installation)
  - [Release History](#release-history)
  - [Versioning](#versioning)
  - [Authors](#authors)
  - [Articles](#articles)
  - [Tools](#tools)

## Modules

The table below contains dependencies to other PowerShell modules.

Module Name | Description
-- | --
Az.Accounts | Needed for the module Az.KeyVault.
Az.KeyVault | To have access to the Azure KeyVault to get the credentials.
MSAL.PS (version 4.36.1.2) | Needed this authentication to send mail.

## Installation

Download the latest version of the PowerShell module [Boehringer.ITEDS.SharePoint.PowerShell](release/Boehringer.ITEDS.SharePoint.PowerShell.zip). Upload this *ZIP* file to *Modules* of your **Automation Account**.

## Release History

Please read [release-notes.md](release-notes.md) for details on getting them.

## Versioning

We use SemVer for versioning.

## Authors

Stefan Gericke - *Initial work* - Boehringer Ingelheim - <stefan.gericke@boehringer-ingelheim.com>


## Articles

- Microsoft Docs: [How to Write a PowerShell Script Module](https://docs.microsoft.com/en-us/powershell/scripting/developer/module/how-to-write-a-powershell-script-module?view=powershell-7.2)

## Tools

- Visual Studio Code

<!-- References -->
[Bitbucket Repo]: https://portal.azure.com/#@e1f8af86-ee95-4718-bd0d-375b37366c83/resource/subscriptions/59f03347-1c98-4abb-bfd8-4f329cf6349d/resourceGroups/eds-ccm-m365-automation-westeurope-prod-5f58710b-rg/

<!-- Shields -->
[Version]: https://img.shields.io/badge/Boehringer.ITEDS.SharePoint.PowerShell-Version%200.1.0-blue?style=flat-square
[Language]: https://img.shields.io/badge/Language-PowerShell-blue?logo=powershell&style=flat-square
[Platform]: https://shields.io/badge/Platform-Azure%20Automation-blue?logo=azurefunctions&style=flat-square
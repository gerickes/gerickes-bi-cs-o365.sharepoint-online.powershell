# Release Notes

Release notes of the Git repository [Bitbucket: bi-cs-o365.sharepoint-online.powershell](https://bitbucket.biscrum.com/projects/SPO/repos/bi-cs-o365.sharepoint-online.powershell/).

## Table of Contents
- [Release Notes](#release-notes)
  - [Table of Contents](#table-of-contents)
  - [Development](#development)
  - [Releases](#releases)
    - [v0.1.0 - December 8th, 2021](#v010---december-8th-2021)

## Development

Commits in Bitbucket of *master* branch: [Commits](https://bitbucket.biscrum.com/projects/SPO/repos/bi-cs-o365.sharepoint-online.powershell/commits)

## Releases 

### v0.1.0 - December 8th, 2021

New:

- Authentication
    - Get-GraphAPIAccessToken: Get *access token* from Graph API by using *Managed Identity*.
    - Get-KeyVaultCredential: Get object `PSCredential` from user stored in variables and password stored in *Azure KeyVault*.
- E-Mail
    - Get-ContentTypeFromFileName: Used for sending attachementsa in a mail.
    - Send-O365MailMessage: Send mail with using an *Enterprise Application* in Azure.
  - Logging
    - New-LogFile: Create new log file.
    - Write-Log: Write log information in the log file.
    - Add-LogFileToSpo: Upload log file to a SharePoint Online site.

[![Download]]((release/Boehringer.ITEDS.SharePoint.PowerShell.zip))

<!-- Shields -->
[Download]: https://img.shields.io/badge/Download-Boehringer.ITEDS.SharePoint.PowerShell-blue?style=flat-square
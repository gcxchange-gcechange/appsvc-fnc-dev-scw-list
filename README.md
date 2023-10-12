# GCX List Operations

## Summary

A collection of function apps for managing the community request SharePoint list:
- Add a community (site) creation request as an item in a SharePoint list
- Update the status of a community request
- Act on a request decision - Approved or Rejected
  
## Prerequisites

The following user accounts (as reflected in the app settings) are required:

| Account             | Membership requirements                                  |
| ------------------- | -------------------------------------------------------- |
| delegatedUserName   | |

## Version 

![dotnet 6](https://img.shields.io/badge/net6.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type        | Admin consent | Justification                             |
| ------------------------- | ----------- | ------------- | ----------------------------------------- |
| Sites.ReadWrite.All       | Delegated   | Yes           | Add/update items in the site request list | 

## App setting

| Name                | Description                                                                       |
| ------------------- | --------------------------------------------------------------------------------- |
| AzureWebJobsStorage | Connection string for the storage acoount                                     	  |
| clientId            | The application (client) ID of the app registration                           	  |
| delegatedUserName   | Delegated authentication user for applying the template and updating Request list |
| delegatedUserSecret | The secret name for the delegated user password                                   |
| keyVaultUrl         | Address for the key vault                                                     	  |
| listId              | Id of the SharePoint list for community requests                                  |
| secretName          | Secret name used to authorize the function app                                	  |
| siteId              | The secret name for the delegated user password                                   |
| tenantId            | Id of the Azure tenant that hosts the function app                                |

## API call parameters (CreateItem)

| Name | Description |
|----- |------------ |
| SecurityCategory | The requested sensitivity category |
| SpaceName | English name of the space |
| SpaceNameFR | French name of the space |
| Owner1 | Owner of the space |
| SpaceDescription | English description of the space |
| SpaceDescriptionFR | French description of the space |
| TemplateTitle | Title of the space template |
| TeamPurpose | Purpose for the existence of the space |
| BusinessJustification | Justification for the existence of the space |
| RequesterName | Name of the person making the request |
| RequesterEmail | Email address of the person making the request |
| Status | Status of the space |

## Version history

Version|Date|Comments
-------|----|--------
1.0|2023-10-10|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

# Organization Information List

## Summary

This function retrive items from 2 sharepoint list. One for domain and the other for department.

## Prerequisites

## Version 

![dotnet 8](https://img.shields.io/badge/net8.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type        | Admin consent | Justification                       |
| ------------------------- | ----------- | ------------- | ----------------------------------- |
|  Sites.Read.All           | Application | Yes           | Read lists in site collection       | 
|  User.Read                | Delegated   | Yes           | Sign in and read user profile       | 

Sharepoint

n/a

## App setting

| Name                     | Description                                                                       |
| ------------------------ | --------------------------------------------------------------------------------- |
| AzureWebJobsStorage      | Connection string for the storage acoount                                         |
| clientId                 | The application (client) Id of the app registration                               |
| keyVaultUrl              | Address for the key vault                                                         |
| ListDepartment           | Id of the sharepoint list department                                              |
| ListDomain               | Id of the sharepoint list domain                                                  |
| secretName               | Secret name used to authorize the function app                                    |
| siteId                   | Id of the site that contains the lists                                            |
| tenantid                 | Id of the Azure tenant that hosts the function app                                |

## Version history

| Version | Date        | Comments         |
| --------| ----------- | ---------------- |
| 1.0     | TBD         | Initial release  |
| Current | 2025-09-09  | Dotnet 8 upgrade |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
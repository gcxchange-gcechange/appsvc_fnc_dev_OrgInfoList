
# Organization Information List

## Summary

This function retrive items from 2 sharepoint list. One for domain and the other for department.

## Prerequisites

n/a

## Version 

![dotnet 10](https://img.shields.io/badge/net10.0-blue.svg)

## API permission

MSGraph

| API / Permissions name    | Type        | Admin consent | Justification                       |
| ------------------------- | ----------- | ------------- | ----------------------------------- |
| Sites.Read.All            | Application | Yes           | Read all sites                      |
| User.Read                 | Delegated   | No            | Sign in and read user profile       |

Sharepoint

n/a

## App setting

| Name                    | Description                                                                    |
| ----------------------- | ------------------------------------------------------------------------------ |
| AzureWebJobsStorage     | Connection string for the storage acoount                                      |
| clientId                | The application (client) ID of the app registration                            |
| keyVaultUrl             | Address for the key vault                                                      |
| ListDepartment          | ID of the SharePoint list for departments                                      |
| ListDomain              | ID of the SharePoint list for domains                                          |
| secretName              | Secret name used to authorize the function app                                 |
| siteId                  | ID of the SharePoint site where the lists are located                          |
| tenantId                | Id of the Azure tenant that hosts the function app                             |

## Version history

Version | Date	     | Comments         |
------- | ---------- | ---------------- |
1.0     | TBD        | Initial release  |
1.1     | 2026-11-26 | dotnet10 upgrade |

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

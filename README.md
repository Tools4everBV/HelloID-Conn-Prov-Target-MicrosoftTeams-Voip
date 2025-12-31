# HelloID-Conn-Prov-Target-MicrosoftTeams-Voip

<!--
** for extra information about alert syntax please refer to [Alerts](https://docs.github.com/en/get-started/writing-on-github/getting-started-with-writing-and-formatting-on-github/basic-writing-and-formatting-syntax#alerts)
-->

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.

<p align="center">
  <img src="">
</p>

## Table of contents

- [HelloID-Conn-Prov-Target-MicrosoftTeams-Voip](#helloid-conn-prov-target-microsoftteams-voip)
  - [Table of contents](#table-of-contents)
  - [Introduction](#introduction)
  - [Available Examples](#available-examples)
  - [Supported  features](#supported--features)
  - [Getting started](#getting-started)
    - [Prerequisites](#prerequisites)
    - [Connection settings](#connection-settings)
    - [Correlation configuration](#correlation-configuration)
    - [Field mapping](#field-mapping)
  - [Remarks](#remarks)
    - [Implementation Complexity and Customization](#implementation-complexity-and-customization)
    - [CallingLineIdentity Mapping](#callinglineidentity-mapping)
    - [Teams Phone License Requirement](#teams-phone-license-requirement)
    - [Enterprise Voice Setting](#enterprise-voice-setting)
    - [Certificate-Based Authentication](#certificate-based-authentication)
    - [No Account Creation Support](#no-account-creation-support)
    - [Authentication Method](#authentication-method)
  - [Development resources](#development-resources)
    - [API endpoints](#api-endpoints)
    - [API documentation](#api-documentation)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Introduction

_HelloID-Conn-Prov-Target-MicrosoftTeams-Voip_ is a _target_ connector that manages Microsoft Teams Direct Routing phone number assignments and calling line identities for users.

This connector enables automatic provisioning and management of Teams telephony features for users within your organization. It correlates existing Microsoft Teams users (based on their UPN or other identifiers) and manages their calling line identity (CallingLineIdentity) based on department information and Enterprise Voice settings.

**Key functionalities:**
- **Correlation-based**: The connector correlates with existing Microsoft Teams users that have a Teams Phone (Phone System/MCOEV) license assigned
- **Calling Line Identity Management**: Automatically assigns and updates the appropriate CallingLineIdentity policy based on the user's department ID
- **Enterprise Voice Control**: Manages the EnterpriseVoiceEnabled setting for users
- **Teams Phone License Validation**: Verifies that users have the required Teams Phone System license before making any changes

**Note:** This connector does not create new Teams users. Users must already exist in Microsoft Teams and have a Teams Phone license assigned through Microsoft 365 license management.

> [!WARNING]
> **This is a complex connector that requires customization for each implementation.** Due to the nature of Microsoft Teams policies (which can only be assigned once to a user and cannot be managed as traditional HelloID permissions), all policy management actions must be implemented directly in the create and update scripts. Each customer has unique requirements for Teams telephony, making customization necessary. The connector in the main branch provides a robust framework for building custom implementations. **Please account for additional implementation time when planning your deployment.**

## Available Examples

Because every implementation requires customization, this repository includes example implementations to help you get started:

- **Main Branch (Recommended Starting Point)**<br>
  The standard connector in the main branch manages CallingLineIdentity policies based on department IDs. This provides a solid framework that you can extend with additional functionality based on your specific requirements.

- **Example-v1 Folder**<br>
  Contains the legacy v1 connector with its own functionality. While this connector is no longer actively maintained and cannot be deployed as-is, it can provide valuable insights and ideas for implementing specific features in the current v2 connector.

- **AdvancedExample Folder**<br>
  Demonstrates a more complex v2 implementation that includes database connectivity for additional data lookups and advanced policy management scenarios.

> [!TIP]
> **Implementation Approach:** We recommend starting with the standard connector from the main branch and extending it based on your organization's specific Teams telephony requirements. This approach ensures you have a solid foundation while allowing for the necessary customization.

## Supported  features

The following features are available:

| Feature                                   | Supported | Actions                                 | Remarks           |
| ----------------------------------------- | --------- | --------------------------------------- | ----------------- |
| **Account Lifecycle**                     | ✅        | Create, Update                          |                   |
| **Permissions**                           | ❌        | -                                       |                   |
| **Resources**                             | ❌        | -                                       |                   |
| **Entitlement Import: Accounts**          | ❌        | -                                       |                   |
| **Entitlement Import: Permissions**       | ❌        | -                                       |                   |
| **Governance Reconciliation Resolutions** | ❌        | -                                       |                   |

## Getting started

### Prerequisites

This connector requires the following prerequisites:

- **Microsoft Teams PowerShell Module**:<br>
  The MicrosoftTeams PowerShell module (version 5.7.1 or higher) must be installed and available on the HelloID agent. Please see the [Microsoft documentation](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install) for installation instructions. The module [can be found here](https://www.powershellgallery.com/packages/MicrosoftTeams). The connector will automatically import the required cmdlets (`Get-CsOnlineUser`, `Get-CsPhoneNumberAssignment`, `Set-CsPhoneNumberAssignment`, `Get-CsCallingLineIdentity`, `Grant-CsCallingLineIdentity`).

- **On-Premises Agent Required**:<br>
  This connector is required to run on an **On-Premises** HelloID agent, as it is not allowed to import PowerShell modules with the Cloud Agent.

- **Microsoft Entra ID App Registration**:<br>
  An App Registration in Microsoft Entra ID is required with the following configuration:
  - Application (client) ID
  - Directory (tenant) ID
  - Client certificate (Base64 encoded PFX/PKCS#12) with private key
  - Certificate password
  
- **App Registration Permissions**:<br>
  The App Registration requires the following Microsoft Graph API permissions as **Application permissions**:
  - **Organization.Read.All** - Read organization information
  
  To grant admin consent to the application, navigate to **Microsoft Entra Portal > Microsoft Entra ID > App Registrations > [Your App] > API Permissions** and press the "**Grant admin consent for [TENANT]**" button.

- **Microsoft Entra ID Role Assignment**:<br>
  The App Registration must be assigned the **Teams Communications Administrator** role (or Global Administrator for more extensive permissions). This role provides the required permissions to manage Teams calling settings. To assign the role:
  - Navigate to **Microsoft Entra Portal > Microsoft Entra ID > Roles and administrators**
  - Select "**Teams Communications Administrator**"
  - Click "**Add assignments**" and select your App Registration
  - Verify the assignment is successful
  
  For more information, see [Application-based authentication in Teams PowerShell Module](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-application-authentication#setup-application-based-authentication) and [Use Microsoft Teams administrator roles](https://learn.microsoft.com/en-us/microsoftteams/using-admin-roles).

- **Teams Phone License**:<br>
  Users must have a Microsoft Teams Phone System license (MCOEV) assigned. The connector validates this before making any changes.

- **CallingLineIdentity Policies**:<br>
  CallingLineIdentity policies must be preconfigured in Microsoft Teams with their Description field matching the department IDs from your source system.

### Connection settings

The following settings are required to connect to Microsoft Teams PowerShell.

| Setting                          | Description                                                                        | Mandatory |
| -------------------------------- | ---------------------------------------------------------------------------------- | --------- |
| TenantID                         | The Azure AD Directory (tenant) ID where the App Registration is located          | Yes       |
| AppId                            | The Application (client) ID of the Azure AD App Registration                       | Yes       |
| AppCertificateBase64String       | The Base64 encoded certificate (PFX/PKCS#12) with private key                     | Yes       |
| AppCertificatePassword           | The password for the certificate                                                   | Yes       |

### Correlation configuration

The correlation configuration is used to specify which properties will be used to match an existing Teams user within _Microsoft Teams_ to a person in _HelloID_.

| Setting                   | Value                             |
| ------------------------- | --------------------------------- |
| Enable correlation        | `True`                            |
| Person correlation field  | `Person.Accounts.MicrosoftActiveDirectory.userPrincipalName` |
| Account correlation field | `UserPrincipalName`                        |

> [!IMPORTANT]
> Correlation is **mandatory** for this connector. The connector does not support creating new Teams users, as Teams users are managed through Microsoft 365 license assignment. The connector will correlate with existing Teams users based on their UserPrincipalName or another identifier.

> [!TIP]
> _For more information on correlation, please refer to our correlation [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems/correlation.html) pages_.

### Field mapping

The field mapping can be imported by using the _fieldMapping.json_ file.

## Remarks

### Implementation Complexity and Customization

This connector is inherently complex due to the nature of Microsoft Teams policy management:

**Why Customization is Required:**
- **Single Policy Assignment**: Teams policies (such as CallingLineIdentity, VoiceRoutingPolicy, etc.) can only be assigned once to a user. Unlike traditional HelloID permissions that can be granted and revoked independently, these policies represent a single configuration state.
- **No Permission Model**: Because policies cannot be managed as separate permissions, all policy logic must be implemented directly in the create and update scripts.
- **Customer-Specific Requirements**: Every organization has unique telephony requirements, routing rules, number ranges, and business logic that need to be reflected in the connector implementation.

**Framework Approach:**

The connector in the main branch provides a framework that handles:
- Microsoft Teams authentication and connection management
- User correlation and validation
- Teams Phone license verification
- Basic CallingLineIdentity policy assignment based on department ID
- Error handling and logging

You can extend this framework by adding:
- Additional policy types (VoiceRoutingPolicy, TenantDialPlan, etc.)
- Phone number assignment logic
- Integration with external systems (databases, other APIs)
- Custom business rules and validation
- Complex routing logic based on multiple attributes

**Implementation Planning:**

When implementing this connector, plan for:
- Requirements gathering session with the customer to understand their Teams telephony setup
- Mapping customer-specific policies to HelloID business rules
- Custom script development and testing
- Extended implementation time compared to standard connectors

### CallingLineIdentity Mapping

The connector maps CallingLineIdentity policies to users based on their department ID. The CallingLineIdentity policies in Microsoft Teams must have their **Description** field set to match the department ID from your source system.

**Example:**
- If a user has `DepartmentId = "DEPT001"` in their HelloID account data
- The connector will look up a CallingLineIdentity policy where `Description = "DEPT001"`
- This CallingLineIdentity will then be assigned to the user

**Important:** 
- Each department ID must have exactly one matching CallingLineIdentity policy
- The connector will throw an error if no matching policy is found or if multiple policies match the same department ID
- Make sure CallingLineIdentity policies are configured in Microsoft Teams before provisioning users

### Teams Phone License Requirement

The connector validates that users have a Teams Phone System license (feature type `PhoneSystem` or provisioned plan `MCOEV`) before making any changes. If a user doesn't have this license:
- During correlation (create action): The connector will throw an error
- During update: The connector will log a warning but continue processing

Users must receive their Teams Phone license through normal Microsoft 365 license management before this connector can manage their calling settings.

### Enterprise Voice Setting

The `EnterpriseVoiceEnabled` property controls whether a user can make and receive calls through Teams Direct Routing. This setting is managed using the `Set-CsPhoneNumberAssignment` cmdlet and can be toggled through the connector's field mapping.

### Certificate-Based Authentication

The connector uses certificate-based authentication for connecting to Microsoft Teams PowerShell. The certificate is:
- Provided as a Base64 encoded string (PFX/PKCS#12 format)
- Temporarily loaded into the CurrentUser certificate store during execution
- Used for authenticating the PowerShell session with the App Registration credentials

### No Account Creation Support

This connector does **not** support creating new Teams user accounts. All users must already exist in Microsoft Teams (through Azure AD sync or other means) and have appropriate licenses assigned. The connector only manages calling-related settings for existing users.

### Authentication Method

This connector uses certificate-based authentication with the Teams PowerShell module. **Client Secret authentication is not supported** - only certificate-based authentication is implemented. Make sure to provide a valid certificate in Base64 encoded format (PFX/PKCS#12) along with the certificate password.

## Development resources

### API endpoints

This connector uses Microsoft Teams PowerShell cmdlets rather than direct API calls. The following cmdlets are used:

| Cmdlet                         | Description                                                                  |
| ------------------------------ | ---------------------------------------------------------------------------- |
| Connect-MicrosoftTeams         | Establishes a connection to Microsoft Teams using certificate authentication |
| Get-CsOnlineUser               | Retrieves Microsoft Teams user information and license details               |
| Get-CsPhoneNumberAssignment    | Gets phone number assignment information for users                           |
| Set-CsPhoneNumberAssignment    | Sets or updates phone number assignments and Enterprise Voice settings       |
| Get-CsCallingLineIdentity      | Retrieves configured CallingLineIdentity policies                            |
| Grant-CsCallingLineIdentity    | Assigns a CallingLineIdentity policy to a user                               |

### API documentation

- [Microsoft Teams PowerShell Module Documentation](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-overview)
- [Set-CsPhoneNumberAssignment](https://docs.microsoft.com/en-us/powershell/module/teams/set-csphonenumberassignment)
- [Grant-CsCallingLineIdentity](https://docs.microsoft.com/en-us/powershell/module/teams/grant-cscallinglineidentity)
- [Get-CsOnlineUser](https://docs.microsoft.com/en-us/powershell/module/teams/get-csonlineuser)

## Getting help

> [!TIP]
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages_.

> [!TIP]
>  _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/

# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [1.0.0] - 2025-12-31

This is the first official release of _HelloID-Conn-Prov-Target-MicrosoftTeams-Voip_. This release is based on template version _v3.1.0_.

### Added

- Comprehensive README documentation describing connector functionality
- Detailed prerequisites section including:
  - Microsoft Teams PowerShell module version requirement (5.7.1 or higher)
  - On-Premises agent requirement
  - Microsoft Entra ID App Registration setup instructions
  - Required API permissions (Organization.Read.All)
  - Microsoft Entra ID role assignment instructions (Teams Communications Administrator)
- Introduction section explaining connector purpose and key functionalities:
  - Correlation-based user management
  - CallingLineIdentity policy management based on department ID
  - Enterprise Voice control
  - Teams Phone license validation
- Implementation complexity section explaining:
  - Why customization is required for each implementation
  - Framework approach and extensibility options
  - Implementation planning considerations
- Available Examples section documenting:
  - Main branch standard connector (recommended starting point)
  - Example-v1 folder with legacy connector for reference
  - AdvancedExample folder with database integration example
- Remarks section with detailed explanations:
  - CallingLineIdentity mapping to department IDs
  - Teams Phone license requirements
  - Enterprise Voice settings management
  - Certificate-based authentication details
  - No account creation support clarification
  - Authentication method limitations (certificate only)
- Development resources section with:
  - List of Microsoft Teams PowerShell cmdlets used
  - Links to Microsoft documentation
- Connection settings documentation for certificate-based authentication
- Correlation configuration with example mapping
- Warning banner about implementation complexity and customization requirements

### Changed

- Updated connector title to reflect Microsoft Teams VoIP/Direct Routing functionality
- Updated correlation configuration to use UserPrincipalName instead of generic EmployeeNumber
- Updated supported features table to show only Create and Update actions (removed Enable, Disable, Delete)
- Updated connection settings to reflect certificate-based authentication requirements (removed generic username/password/baseurl)

### Deprecated

### Removed

- Generic placeholder text from template
- Unused API endpoints section (replaced with PowerShell cmdlets)
- Template example remarks that weren't applicable
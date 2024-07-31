# Set-NinjaUserFromOST PowerShell Scripts

This repository contains two versions of a PowerShell script designed to detect the most recently active user on a Windows machine and set their User Principal Name (UPN) as a custom field in NinjaRMM to later be used as an import matching anchor for HaloPSA. This will allow you to correctly match devices to users in HaloPSA during import from NinjaRMM (both manually and during the regular integrator runs). This dramatically improves the accuracy and success rate of user matching during import.

## Script Versions

### 1. Without Graph API Integration

- [README for Without Graph Version](Without_Graph\README_WithoutGraph.md)
- [Script: Set-NinjaUserFromOST.ps1](Without_Graph\Set-NinjaUserFromOST.ps1)

This version of the script operates locally on the Windows machine without requiring any Azure AD or Microsoft Graph API integration. It's simpler to set up and use but may not provide the same level of verification of validity for user accounts.

### 2. With Graph API Integration

- [README for Graph-Enabled Version](With_Graph\README_GraphEnabled.md)
- [Script: Set-NinjaUserFromOST_GraphEnabled.ps1](With_Graph\Set-NinjaUserFromOST_GraphEnabled.ps1)

This version integrates with Microsoft Graph API to verify user accounts against Azure AD. It provides an additional layer of accuracy but requires more setup, including an Azure AD app registration and appropriate permissions.

## Key Differences

1. **Authentication**: The Graph-enabled version requires Azure AD app registration and authentication.
2. **User Verification**: The Graph-enabled version verifies users against Azure AD, while the other version relies solely on local data.
3. **Setup Complexity**: The Graph-enabled version requires more initial setup but may provide more accurate results in multi-tenant environments.
4. **Dependencies**: The Graph-enabled version has additional dependencies on Azure AD and Microsoft Graph API.

Choose the version that best fits your environment and requirements. If you're operating in a multi-tenant Azure AD environment and need additional verification, use the Graph-enabled version. For simpler setups or environments, the version without Graph integration may be more suitable.

## Contributing

Contributions to either version of the script are welcome. Please submit a pull request or open an issue to discuss proposed changes.

## License

Copyright (c) 2024 TechPulse Consulting LLC and Christopher Scaminaci

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
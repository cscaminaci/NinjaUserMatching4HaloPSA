# SetNinjaUserCF PowerShell Script

This PowerShell script automates the process of detecting the most recently active user on a Windows machine and setting their User Principal Name (UPN) as a custom field in NinjaRMM. It's designed to work with multi-tenant Azure AD environments and integrates with Microsoft Graph API.

## How It Works

1. Scans all user profiles on the machine.
2. Analyzes Outlook OST files to extract valid company UPNs.
3. Determines the most recently active user based on OST and profile last write times.
4. Authenticates with Azure AD using a multi-tenant app registration.
5. Verifies the user's existence in Azure AD using Microsoft Graph API.
6. Sets the user's UPN as a custom field in NinjaRMM.

## Prerequisites

- Windows machine with PowerShell 5.1 or later
- NinjaRMM agent installed on the target machine
- Multi-tenant Azure AD app registration with the following:
  - Client ID
  - Client Secret
  - Tenant ID (source tenant)
- Microsoft Graph API permissions for the app:
  - User.Read.All (Application permission)
- Outlook installed on the target machine
- Custom field created in NinjaRMM named "lastDetectedUser"

## Setup for NinjaRMM

1. Create a custom field in NinjaRMM named "lastDetectedUser" (you can use anything else you'd like, just update the script to match).
2. Set the custom field to have Read:API permission
3. Upload the script to your NinjaRMM script library.
4. Replace the placeholders in the script with your actual values:
   - `<multitenant app id>`: Your Azure AD app Client ID
   - `<multitenant app secret>`: Your Azure AD app Client Secret
   - `<app source tenant>`: Your Azure AD Tenant ID

## Usage with NinjaRMM

1. Create a new automation on your policy in NinjaRMM.
2. Select the uploaded script.
3. Set the execution schedule as needed (I recommend on user login).
4. Set the run as to SYSTEM

## Setup for Halo

1. Create an asset field on your NinjaRMM asset types called "Last Detected User" or whatever you'd like to name it.
2. In the NinjaRMM field mappings, set the new field to map to "*Ninja Custom Field*" and type in the name of the custom field you made in Ninja.
3. Enable to checkbox to attempt matching users on import
4. Type the name of your Ninja custom field into the text box for Ninja Custom Field to attempt matching user on.

## Error Handling

If the script encounters any errors, it will:
1. Write the error message to the standard error stream.
2. Attempt to set the "lastDetectedUser" custom field to the error message.
3. Exit with a non-zero status code.

## Notes

- Ensure that the multi-tenant app has the necessary permissions in all relevant tenants.
- The script filters out common free email domains to focus on company UPNs.

## Contributing

Contributions to improve the script are welcome. Please submit a pull request or open an issue to discuss proposed changes.

## License

[MIT License](LICENSE)
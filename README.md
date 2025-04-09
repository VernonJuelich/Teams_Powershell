# Custom PowerShell Scripts for Supporting Teams Calling Environments

## ContactImport.ps1
This PowerShell script leverages the **AADInternals** module to manage contacts effectively for Teams Calling environments.
### Features
1. **Targeted User UPN Prompt**
   - Prompts for a **User** or **CAP account**.
2. **Contact Type Selection**
   - Asks whether the contact is **internal** or **external**:
     - **Internal Contact**:
       - Prompts for the contact's email address.
       - Retrieves contact details from **Entra ID**.
     - **External Contact**:
       - Prompts for the contact's **First Name**, **Last Name**, and **Phone Number**.
3. **Speed Dial Configuration**
   - Provides an option to set the contact as a **Speed Dial**.
---

# PowerShell Script for Managing Microsoft Teams Voice Public Holiday Schedules

## AU Teams Holiday Mangement.ps1
This PowerShell script streamlines the management of Public Holiday schedules within your Microsoft Teams Voice environment. It fetches the latest public holiday data for specified Australian states from data.gov.au and provides options to either create new holiday schedules or update existing ones.
### Features
1. **Australian State Selection**
   - Prompts you to enter the Australian state codes (e.g., sa, qld, nsw) for which you want to manage public holidays.
2. **Data Source Integration**
   - Retrieves up-to-date public holiday information directly from data.gov.au.
3. **New Schedule Creation** (Also done via script AU New Public Holidays.ps1)
   - Allows you to create new Teams Voice holiday schedules based on the fetched data.
   - Offers the flexibility to apply the same action to all imported holidays or configure actions individually.
4. **Existing Schedule Updates** (Also done via script AU Update Public Holidays.ps1)
   - Identifies and updates existing Teams Voice holiday schedules with the latest dates.
   - Updates are intelligently applied only to holidays falling within the current or the following year.
   - Attempts to match holidays based on specific naming conventions or broader state-level groupings.
### Prerequisites
- **Microsoft Teams PowerShell Module:** Ensure it's installed (`Install-Module MicrosoftTeams`).
- **Internet Connectivity:** The script needs to access data.gov.au.
- **Teams Voice Permissions:** You'll need sufficient permissions to manage call flow schedules in Microsoft Teams.
### Assumptions
- **Initial Schedule Creation:** This script assumes that if you're using the 'update' mode, the initial holiday schedules were likely created using a similar naming convention (e.g., "NSW Boxing Day" or "NSW Public Holiday").
### Notes
- **Skipping Specific Holidays:** To exclude specific holidays during the import, you can modify the `Get-PublicHolidays` function around line 83 to filter by the 'Holiday Name'.
---

## AACQ RA Creation.ps1
This PowerShell script streamlines the creation of Microsoft Teams Resource Accounts for Auto Attendants and Call Queues by processing data from CSV files. It also offers optional licensing and Usage Location configuration.
### Features
1. **Bulk Creation of Resource Accounts**
   - Creates multiple Auto Attendant and Call Queue accounts based on CSV data.
2. **CSV-Driven Input**
   - Reads account details (UPN, Display Name) from simple CSV files.
3. **Optional Licensing**
   - Can automatically assign the 'Microsoft Teams Phone Resource Account' license.
4. **Usage Location Configuration**
   - Automatically sets the Usage Location to 'AU' for licensed accounts.
5. **Separate Processing for AA and CQ**
   - Handles Auto Attendant and Call Queue accounts independently.
---


Feel free to let me know if you'd like any further adjustments or additions! 😊

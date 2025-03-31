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

Feel free to let me know if you'd like any further adjustments or additions! 😊

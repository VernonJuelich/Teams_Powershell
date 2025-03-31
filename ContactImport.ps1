# Import the Azure AD Internals module.
#Install-Module AADInternals
#Import-Module AADInternals -Force -ErrorAction Stop

# Define the base URL for the Teams contacts API.
$BaseURL = "https://teams.microsoft.com/api/mt/part/au-01/beta/contactsv3/"

try {
    # Prompt for username and password.  Do this *once* at the beginning.
    $username = Read-Host -Prompt "Enter target accounts username (UPN)"
    $password = Read-Host -AsSecureString -Prompt "Enter target accounts password"
    $credential = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)

    # Get the Teams token for the user.  Get this *once*
    $token = Get-AADIntAccessTokenForTeams -Credentials $credential
    $adtoken = Get-AADIntAccessTokenForAADGraph -Credentials $credential  -SaveToCache

    # Define the headers for the API request.
    $header = @{ Authorization = "Bearer $token" }

    # Prompt for contact type (internal or external)
    $contactType = Read-Host -Prompt "Is the contact internal or external? (internal/external)"

    if ($contactType -eq "internal") {
        # Prompt for internal contact details
        $email = Read-Host -Prompt "Enter the email address of the internal contact"
        $speedDial = Read-Host -Prompt "Add to speed dial? (Y/N)"

        # Initialize variables
        $aadUser = $null
        $businessNumber = ""
        $userObjectId = ""
        $tenantId = ""  # Initialize tenant ID
        $jobTitle = ""    # Default value
        $displayName = "" # Default
        $officeLocation = "" # Default
        $phoneType = "Business"
        $systemTag = ""
        $systemTagId = ""

        try {
            # Attempt to retrieve user data from Azure AD
            $aadUser = Get-AADIntUser -UserPrincipalName $email -ErrorAction Stop

            if ($aadUser) {
                # Extract user information from Azure AD
                $first = $aadUser.FirstName
                $last = $aadUser.LastName
                $businessNumber = $aadUser.PhoneNumber
                $displayName = "$first $last"
                $userObjectId = $aadUser.ObjectId  # Get the User Object ID.
                $tenantId = Get-AADIntTenantID -Domain ($email.Split('@')[1]) # Get Tenant ID
                $mri = "8:orgid:$userObjectId" #format mri
                $jobTitle = $aadUser.Title # Get job title from AAD
                $officeLocation = $aadUser.City # Get Office Location from AAD
            } else {
                Write-Warning "User not found in Azure AD for email: $email."
                # Handle the case where the user is not found in AAD.  You might
                # choose to exit, or prompt for manual entry of details.  For this
                # example, we'll exit.
                return  # Exit the script.  You could also use "exit" here.
            }
        } catch {
            Write-Warning "Could not retrieve user information from Azure AD for email: $email. $($_.Exception.Message)"
            return # Exit.
        }

        # Set system tag based on Speed Dial
        if ($speedDial -eq "Y") {
            $systemTag = "Favorites"
            $systemTagId = "systemtags_Favorites"
        }

        # Construct the JSON payload for internal contact.
        $payload = @"
{
    "mri": "$mri",
    "defaultEmail": "$email",
    "names": [
        {
            "displayName": "$displayName",
            "first": "$first",
            "last": "$last",
            "middle": "",
            "nickname": "",
            "suffix": "",
            "title": "",
            "pronunciation": {
                "first": "",
                "last": ""
            },
            "source": {
                "type": "Profile"
            }
        }
    ],
    "emails": [
        {
            "address": "$email",
            "type": "Main",
            "source": {
                "type": "Profile"
            }
        }
    ],
    "phones": [
        {
            "number": "$businessNumber",
            "type": "$phoneType",
            "source": {
                "type": "Profile"
            }
        }
    ],
    "addresses": [],
    "positions": [
        {
            "detail": {
                "jobTitle": "$jobTitle",
                "company": {
                    "displayName": " ",
                    "pronunciation": "",
                    "department": "",
                    "officeLocation": "$officeLocation"
                }
            },
            "source": {
                "type": "Profile"
            }
        }
    ],
    "notes": [],
    "anniversaries": [],
    "relationships": [],
    "websites": [],
    "webAccounts": [
        {
            "userId": "$email",
            "source": {
                "type": "Profile"
            }
        },
        {
            "userId": "OID:$userObjectId@$tenantId",
            "service": {
                "name": "Profile"
            },
            "source": {
                "type": "Profile"
            }
        }
    ],
    "tags": [],
    "systemTags": [
        {
            "displayName": "$systemTag",
            "id": "$systemTagId"
        }
    ]
}
"@
    } elseif ($contactType -eq "external") {
        # Prompt for external contact details
        $first = Read-Host -Prompt "Enter the first name of the external contact"
        $last = Read-Host -Prompt "Enter the last name of the external contact"
        $businessNumber = Read-Host -Prompt "Enter the phone number of the external contact"
        $speedDial = Read-Host -Prompt "Add to speed dial? (Y/N)"
        $displayName = "$first $last"
        $phoneType = "Business"
        $mri = "4:$businessNumber"
        $systemTag = ""
        $systemTagId = ""

         # Set system tag based on Speed Dial
        if ($speedDial -eq "Y") {
            $systemTag = "Favorites"
            $systemTagId = "systemtags_Favorites"
        }

        # Construct the JSON payload for external contact.
        $payload = @"
{
    "mri": "$mri",
    "systemTags": [
        {
            "displayName": "$systemTag",
            "id": "$systemTagId"
        }
    ],
    "names": [
        {
            "first": "$first",
            "last": "$last",
            "displayName": "$displayName",
            "source": {
                "type": "UserProvided"
            },
            "action": "Post"
        }
    ],
    "phones": [
        {
            "number": "$businessNumber",
            "type": "$phoneType",
            "source": {
                "type": "UserProvided"
            }
        }
    ]
}
"@
    } else {
        Write-Warning "Invalid contact type. Please enter 'internal' or 'external'."
        return  # Exit the script
    }

    # URL for adding contacts
    $AddContactURL = "$BaseURL"

    # Add contact to Teams.
    try {
        $response = Invoke-RestMethod -Uri $AddContactURL -Body $payload -Method POST -Headers $header -ContentType "application/json"
        if ($response) {
            Write-Host "Successfully added contact $($displayName) for $($username)"
        } else {
            Write-Warning "Failed to add contact $($displayName) for $($username). Empty Response"
        }
    } catch {
        Write-Warning "Failed to add contact $($displayName) for $($username). Error: $($_.Exception.Message)"
    }
} catch {
    Write-Error "An error occurred during initial setup: $($_.Exception.Message)"
    Write-Verbose "Full Exception: $_"
} finally {
    Clear-AADIntCache
}

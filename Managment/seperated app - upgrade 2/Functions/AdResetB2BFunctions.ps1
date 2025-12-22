# AdResetB2BFunctions.ps1
# =========================
# CONSTANTS
# =========================
$ResetInvitePath   = "\\share\EAMO\B2B\Reset\ResetInvites.csv"
$NewUserInvitePath = "\\share\EAMO\B2B\NewUsers.csv"
$RedirectUrl       = "https://myapplications.microsoft.com"
$DefaultPassword   = "Password8"
$InternalGroup     = "Office365Users"

# Function to check if user is internal
function Test-InternalUser {
    param(
        [string]$UserInput
    )
    
    $outputTextBox.AppendText("Checking if user is internal...`n")
    
    try {
        # Try to get AD user
        $adUser = $null
        
        # Check by SamAccountName (LAN ID)
        if ($UserInput -notlike "*@*") {
            $adUser = Get-ADUser -Filter "SamAccountName -eq '$UserInput'" `
                -Properties EmployeeID, EmployeeType, MemberOf -ErrorAction SilentlyContinue
        }
        
        # If not found by SamAccountName, try by UserPrincipalName
        if (-not $adUser) {
            $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$UserInput'" `
                -Properties EmployeeID, EmployeeType, MemberOf -ErrorAction SilentlyContinue
        }
        
        # If still not found, try by email
        if (-not $adUser) {
            $adUser = Get-ADUser -Filter "EmailAddress -eq '$UserInput'" `
                -Properties EmployeeID, EmployeeType, MemberOf -ErrorAction SilentlyContinue
        }
        
        if (-not $adUser) {
            $outputTextBox.AppendText("User not found in Active Directory. Assuming external.`n")
            return $false
        }
        
        $outputTextBox.AppendText("User found: $($adUser.SamAccountName) | DisplayName: $($adUser.Name)`n")
        
        # TEST 1: Check if they have an EIN (Employee ID Number)
        $hasEIN = $false
        if (-not [string]::IsNullOrEmpty($adUser.EmployeeID)) {
            $hasEIN = $true
            $outputTextBox.AppendText("[YES] Has EIN: $($adUser.EmployeeID)`n")
        } else {
            $outputTextBox.AppendText("[NO] No EIN found`n")
        }
        
        # TEST 2: Check if EmployeeType is 'E'
        $hasEmployeeTypeE = $false
        if (-not [string]::IsNullOrEmpty($adUser.EmployeeType) -and $adUser.EmployeeType -eq 'E') {
            $hasEmployeeTypeE = $true
            $outputTextBox.AppendText("[YES] EmployeeType is 'E'`n")
        } else {
            $outputTextBox.AppendText("[NO] EmployeeType is not 'E' (found: '$($adUser.EmployeeType)')`n")
        }
        
        # TEST 3: Check if they're a member of Office365Users group
        $isInInternalGroup = $false
        try {
            # Check if group exists
            $group = Get-ADGroup -Filter "Name -eq '$InternalGroup'" -ErrorAction SilentlyContinue
            if ($group) {
                $groupMembers = Get-ADGroupMember -Identity $InternalGroup -ErrorAction SilentlyContinue
                $isInInternalGroup = $groupMembers | Where-Object { $_.SamAccountName -eq $adUser.SamAccountName } | Measure-Object | Select-Object -ExpandProperty Count -gt 0
                
                if ($isInInternalGroup) {
                    $outputTextBox.AppendText("[YES] Member of $InternalGroup group`n")
                } else {
                    $outputTextBox.AppendText("[NO] Not a member of $InternalGroup group`n")
                }
            } else {
                $outputTextBox.AppendText("[INFO] $InternalGroup group not found in AD`n")
            }
        } catch {
            $outputTextBox.AppendText("[INFO] Could not check group membership: $_`n")
        }
        
        # Determine if user is internal
        # User is INTERNAL if ANY of these conditions are true:
        $isInternal = $hasEIN -or $hasEmployeeTypeE -or $isInInternalGroup
        
        $outputTextBox.AppendText("=" * 60 + "`n")
        if ($isInternal) {
            $outputTextBox.AppendText("RESULT: User is INTERNAL (DO NOT SEND INVITE)`n")
            $reasons = @()
            if ($hasEIN) { $reasons += "Has EIN" }
            if ($hasEmployeeTypeE) { $reasons += "EmployeeType='E'" }
            if ($isInInternalGroup) { $reasons += "Member of $InternalGroup" }
            $outputTextBox.AppendText("REASON: " + ($reasons -join ", ") + "`n")
        } else {
            $outputTextBox.AppendText("RESULT: User is EXTERNAL (Safe to invite)`n")
        }
        $outputTextBox.AppendText("=" * 60 + "`n")
        
        return $isInternal
        
    } catch {
        $outputTextBox.AppendText("ERROR checking internal status: $_`n")
        # If we can't determine, assume internal to be safe
        $outputTextBox.AppendText("Assuming internal to prevent accidental invites`n")
        return $true
    }
}

# Function to format phone number
function Format-PhoneNumber {
    param (
        [string]$PhoneNumber
    )
    # Remove non-numeric characters except plus
    $formattedNumber = $PhoneNumber -replace '[^\d\+]', ''
    if (-not $formattedNumber.StartsWith("+1")) {
        $formattedNumber = "+1" + $formattedNumber.TrimStart('+', '1')
    }
    return $formattedNumber
}

# Function to get Azure user smartly
function Get-AzureUserSmart {
    param(
        [string]$InputValue
    )

    try {
        $mgUser = $null
        if ($InputValue -like "*@*") {
            try { 
                $mgUser = Get-MgUser -UserId $InputValue -ErrorAction Stop
                $outputTextBox.AppendText("Azure AD: Found user '$InputValue'`n")
            } 
            catch {
                if ($_.Exception.Message -match "Request_ResourceNotFound" -or $_.Exception.Message -match "404") {
                    $outputTextBox.AppendText("Azure AD: User with UPN '$InputValue' not found.`n")
                } else {
                    $outputTextBox.AppendText("Azure AD: Error querying UPN '$InputValue'. Error: $_`n")
                }
                return $null
            }
        } else {
            try {
                $allUsers = Get-MgUser -All -ErrorAction Stop
                $mgUser = $allUsers | Where-Object { $_.OnPremisesSamAccountName -eq $InputValue } | Select-Object -First 1
                if ($mgUser) {
                    $outputTextBox.AppendText("Azure AD: Found LAN ID '$InputValue' -> UPN: $($mgUser.UserPrincipalName)`n")
                } else {
                    $outputTextBox.AppendText("Azure AD: LAN ID '$InputValue' not found in Azure AD`n")
                }
            } catch { 
                $outputTextBox.AppendText("Error searching Azure AD for LAN ID: $_`n")
                return $null
            }
        }
        return $mgUser
    }
    catch {
        $outputTextBox.AppendText("ERROR in Get-AzureUserSmart: $_`n")
        return $null
    }
}

# Function to log messages
function Log-Message {
    param(
        [string]$Message
    )
    $outputTextBox.AppendText("$Message`n")
}

# Event handler for Reset Password button
if ($adResetPasswordButton) {
    $adResetPasswordButton.Add_Click({
        $outputTextBox.Clear()
        $outputTextBox.AppendText("Starting password reset...`n")
        $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

        $inputValue = $adResetB2BUserTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($inputValue)) {
            $outputTextBox.AppendText("ERROR: Enter LAN ID or UPN`n")
            return
        }

        try {
            # Import Active Directory module if not already loaded
            Import-Module ActiveDirectory -ErrorAction Stop
            
            $adUser = Get-ADUser -Filter { SamAccountName -eq $inputValue -or UserPrincipalName -eq $inputValue } -Properties Enabled -ErrorAction Stop
            $securePassword = ConvertTo-SecureString $DefaultPassword -AsPlainText -Force
            Set-ADAccountPassword -Identity $adUser -Reset -NewPassword $securePassword

            if (-not $adUser.Enabled) {
                Enable-ADAccount $adUser
                $outputTextBox.AppendText("SUCCESS: Password reset to $DefaultPassword | Account was DISABLED, now ENABLED`n")
            } else {
                $outputTextBox.AppendText("SUCCESS: Password reset to $DefaultPassword | Account was already ENABLED`n")
            }
        }
        catch {
            if ($_.Exception.Message -match "Cannot find an object") {
                $outputTextBox.AppendText("ERROR: User '$inputValue' not found in Active Directory`n")
            } else {
                $outputTextBox.AppendText("ERROR resetting password: $_`n")
            }
        }
    })
}

# Event handler for Invite Guest button
if ($adResetGuestButton) {
    $adResetGuestButton.Add_Click({
        $outputTextBox.Clear()
        $outputTextBox.AppendText("Starting guest invitation...`n")
        $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

        $inputValue = $adResetB2BUserTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($inputValue)) {
            $outputTextBox.AppendText("ERROR: Enter LAN ID or UPN`n")
            return
        }

        # FIRST: Check if user is internal
        $isInternal = Test-InternalUser -UserInput $inputValue
        
        if ($isInternal) {
            $outputTextBox.AppendText("=" * 60 + "`n")
            $outputTextBox.AppendText("ACTION BLOCKED: User is INTERNAL`n")
            $outputTextBox.AppendText("DO NOT send guest invite to internal users`n")
            $outputTextBox.AppendText("They will lose access to all internal resources`n")
            $outputTextBox.AppendText("=" * 60 + "`n")
            return
        }
        
        $outputTextBox.AppendText("User confirmed as external. Proceeding with guest invitation...`n")

        # Initialize with basic info from the textbox
        $exportUPN = $inputValue # Default to the input
        $exportID = "" # Will remain blank if user not found

        try {
            # Connect to Microsoft Graph if not already connected
            if (-not (Get-MgContext)) {
                Connect-MgGraph -Scopes "User.Read.All"
            }
            
            $mgUser = Get-AzureUserSmart $inputValue
            if ($mgUser) {
                $exportUPN = $mgUser.UserPrincipalName
                $exportID = $mgUser.Id
            } else {
                $outputTextBox.AppendText("INFO: User not found in Azure AD. Using manual input '$inputValue' for CSV.`n")
            }
        }
        catch {
            $outputTextBox.AppendText("WARNING: Error during Azure AD lookup. Proceeding with manual input: $_`n")
        }

        # Create CSV directory if it doesn't exist
        $csvDir = Split-Path $ResetInvitePath -Parent
        if (-not (Test-Path $csvDir)) {
            New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
            $outputTextBox.AppendText("INFO: Created directory: $csvDir`n")
        }

        # Check if file exists to determine if we need headers
        $needHeaders = -not (Test-Path $ResetInvitePath)
        
        # Prepare the object with the data we have (from Graph or the manual input)
        $exportObject = [PSCustomObject]@{
            UPN    = $exportUPN
            UserID = $exportID
            Url    = $RedirectUrl
        }

        # Export to CSV
        try {
            $exportObject | Export-Csv $ResetInvitePath -Append -NoTypeInformation
            $outputTextBox.AppendText("SUCCESS: Guest invite queued for '$inputValue'`n")
            $outputTextBox.AppendText("       -> Saved to CSV: UPN='$exportUPN', UserID='$exportID'`n")
            
            # Show a preview of what was saved
            if ($needHeaders) {
                $outputTextBox.AppendText("INFO: Created new CSV file with headers`n")
            }
        }
        catch {
            $outputTextBox.AppendText("ERROR saving to CSV '$ResetInvitePath': $_`n")
        }
    })
}

# Event handler for Invite Member button
if ($adResetMemberButton) {
    $adResetMemberButton.Add_Click({
        $outputTextBox.Clear()
        $outputTextBox.AppendText("Starting member invitation...`n")
        $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

        $inputValue = $adResetB2BUserTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($inputValue)) {
            $outputTextBox.AppendText("ERROR: Enter LAN ID or UPN`n")
            return
        }

        # FIRST: Check if user is internal
        $isInternal = Test-InternalUser -UserInput $inputValue
        
        if ($isInternal) {
            $outputTextBox.AppendText("=" * 60 + "`n")
            $outputTextBox.AppendText("ACTION BLOCKED: User is INTERNAL`n")
            $outputTextBox.AppendText("DO NOT send member invite to internal users`n")
            $outputTextBox.AppendText("They will lose access to all internal resources`n")
            $outputTextBox.AppendText("=" * 60 + "`n")
            return
        }
        
        $outputTextBox.AppendText("User confirmed as external. Proceeding with member invitation...`n")

        # Initialize
        $exportEmail = $inputValue # Default to input
        $originalInput = $inputValue
        
        try {
            # Connect to Microsoft Graph if not already connected
            if (-not (Get-MgContext)) {
                Connect-MgGraph -Scopes "User.Read.All"
            }
            
            $mgUser = Get-AzureUserSmart $inputValue
            if ($mgUser) {
                $exportEmail = if ($mgUser.Mail) { $mgUser.Mail } else { $mgUser.UserPrincipalName }
            } else {
                $outputTextBox.AppendText("INFO: User not found in Azure AD. Using input '$inputValue' as email for CSV.`n")
            }
        }
        catch {
            $outputTextBox.AppendText("WARNING: Error during Azure AD lookup. Proceeding with manual input: $_`n")
        }

        # Create CSV directory if it doesn't exist
        $csvDir = Split-Path $NewUserInvitePath -Parent
        if (-not (Test-Path $csvDir)) {
            New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
            $outputTextBox.AppendText("INFO: Created directory: $csvDir`n")
        }

        # Check if file exists to determine if we need headers
        $needHeaders = -not (Test-Path $NewUserInvitePath)
        
        # Prepare the object
        $exportObject = [PSCustomObject]@{
            User              = $originalInput
            UserInvitedEmail  = $exportEmail
            InviteRedirectURL = $RedirectUrl
        }

        # Export to CSV
        try {
            $exportObject | Export-Csv $NewUserInvitePath -Append -NoTypeInformation
            $outputTextBox.AppendText("SUCCESS: Member invite queued for '$originalInput'`n")
            $outputTextBox.AppendText("       -> Saved to CSV: Email='$exportEmail'`n")
            
            if ($needHeaders) {
                $outputTextBox.AppendText("INFO: Created new CSV file with headers`n")
            }
        }
        catch {
            $outputTextBox.AppendText("ERROR saving to CSV '$NewUserInvitePath': $_`n")
        }
    })
}

# Function to check and create CSV directories
function Initialize-CSVDirectories {
    $resetCSVDir = Split-Path $ResetInvitePath -Parent
    $newUserCSVDir = Split-Path $NewUserInvitePath -Parent
    
    if (-not (Test-Path $resetCSVDir)) {
        New-Item -ItemType Directory -Path $resetCSVDir -Force | Out-Null
    }
    
    if (-not (Test-Path $newUserCSVDir)) {
        New-Item -ItemType Directory -Path $newUserCSVDir -Force | Out-Null
    }
}

# Initialize directories when the panel loads
if ($adResetB2BButton) {
    $adResetB2BButton.Add_Click({
        # This runs after the visibility handler in mainlauncher.ps1
        Initialize-CSVDirectories
    })
}

# Clean up function to disconnect from Graph
function Disconnect-GraphSession {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    } catch {
        # Ignore errors if already disconnected
    }
}

# Register cleanup when the window closes
if ($window) {
    $window.Add_Closed({
        Disconnect-GraphSession
    })
}
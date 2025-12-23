# Create User Functions
# =========================
# CONSTANTS
# =========================
$DefaultOU = "OU=Created accounts,OU=WahidTest,OU=Wahid,OU=NewUserStaging-EAMO,DC=windows,DC=nyc,DC=hra,DC=nycnet"
$InternetGroup = "gs-dssallowedinternetusers"
$DefaultPassword = Set

# Function to check for existing users (SIMPLIFIED and FAST)
function Check-ForDuplicateUsers {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$ProposedLanID
    )
    
    $duplicates = @()
    
    try {
        # Check 1: Search for users with same first and last name (EXACT match)
        $sameNameUsers = Get-ADUser -Filter "GivenName -eq '$FirstName' -and Surname -eq '$LastName'" `
            -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled -ErrorAction SilentlyContinue
        
        if ($sameNameUsers) {
            foreach ($user in $sameNameUsers) {
                $duplicates += [PSCustomObject]@{
                    Type = "Exact Name Match"
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    UserPrincipalName = $user.UserPrincipalName
                    Enabled = $user.Enabled
                }
            }
        }
        
        # Check 2: Check if proposed LAN ID already exists
        $existingLanID = Get-ADUser -Filter "SamAccountName -eq '$ProposedLanID'" `
            -Properties DisplayName, SamAccountName, UserPrincipalName, Enabled -ErrorAction SilentlyContinue
        
        if ($existingLanID) {
            $duplicates += [PSCustomObject]@{
                Type = "Duplicate LAN ID"
                DisplayName = $existingLanID.DisplayName
                SamAccountName = $existingLanID.SamAccountName
                UserPrincipalName = $existingLanID.UserPrincipalName
                Enabled = $existingLanID.Enabled
            }
        }
    }
    catch {
        # If search fails, just log and continue
        $global:outputTextBox.AppendText("Note: Could not complete duplicate check: $_`n")
    }
    
    return $duplicates
}

# Function to generate LAN ID based on your rules
function Generate-LanID {
    param(
        [string]$LastName,
        [string]$ManualLanID
    )
    
    # If user provided manual LAN ID, use it
    if (-not [string]::IsNullOrEmpty($ManualLanID)) {
        return $ManualLanID.Trim().ToLower()
    }
    
    # Generate LAN ID: First 4 letters of last name + month (2 digits) + day (2 digits)
    $lastNameClean = $LastName -replace '[^a-zA-Z]', ''  # Remove non-letters
    $lastNameClean = $lastNameClean.ToLower()
    
    # Get first 4 characters of last name (pad if shorter)
    if ($lastNameClean.Length -ge 4) {
        $namePart = $lastNameClean.Substring(0, 4)
    } else {
        $namePart = $lastNameClean.PadRight(4, 'x')
    }
    
    # Get current month and day (MMDD format)
    $currentDate = Get-Date
    $month = $currentDate.Month.ToString("00")  # 2 digits
    $day = $currentDate.Day.ToString("00")      # 2 digits
    $datePart = "$month$day"  # MMDD format
    
    # Combine: namePart + datePart (e.g., huss1218 for Hussain on Dec 18)
    $baseLanID = "$namePart$datePart"
    
    # Check if this LAN ID exists, if so, go DOWN in numbers
    $currentID = $baseLanID
    $counter = 0
    
    while (Get-ADUser -Filter "SamAccountName -eq '$currentID'" -ErrorAction SilentlyContinue) {
        $counter++
        
        # Go DOWN: 1218 → 1217 → 1216, etc.
        # We need to decrement the DATE part (MMDD)
        $dateNumber = [int]$datePart
        
        # Subtract the counter from the date number
        $newDateNumber = $dateNumber - $counter
        
        # If we go below 0101 (Jan 1st), wrap to 1231 (Dec 31st)
        if ($newDateNumber -lt 101) {
            # Start from end of year and work backwards
            $newDateNumber = 1231 - ($counter - ($dateNumber - 101))
        }
        
        # Ensure we have 4 digits
        $newDatePart = $newDateNumber.ToString("0000")
        
        $currentID = "$namePart$newDatePart"
        
        # Safety break
        if ($counter -gt 100) {
            $global:outputTextBox.AppendText("Warning: Could not find unique LAN ID after 100 attempts`n")
            # Fallback: Add random number
            $random = Get-Random -Minimum 1000 -Maximum 9999
            $currentID = "$namePart$random"
            break
        }
    }
    
    return $currentID
}

# Function to handle CSV loading
function Import-UserFromCSV {
    $global:outputTextBox.AppendText("Loading CSV file...`n")
    
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv"
    
    if ($dialog.ShowDialog() -ne "OK") { 
        $global:outputTextBox.AppendText("CSV load cancelled.`n")
        return 
    }
    
    try {
        $lines = Get-Content $dialog.FileName
        if ($lines.Count -lt 2) { 
            $global:outputTextBox.AppendText("Error: CSV has no data rows.`n")
            return 
        }
        
        # Split CSV safely (handles quoted commas)
        $headers = $lines[0] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
        $values = $lines[1] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
        
        # Normalize headers
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $headers[$i] = $headers[$i].Trim('"').ToLower()
            if ($i -lt $values.Count) {
                $values[$i] = $values[$i].Trim('"')
            }
        }
        
        function Get-Val($headerFragment) {
            for ($i = 0; $i -lt $headers.Count; $i++) {
                if ($headers[$i] -like "*$headerFragment*") {
                    return $values[$i]
                }
            }
            return ""
        }
        
        # Populate fields
        $global:createUserFirstNameTextBox.Text = Get-Val "first name"
        $global:createUserLastNameTextBox.Text = Get-Val "last name"
        $global:createUserEmailTextBox.Text = Get-Val "email"
        $global:createUserPhoneTextBox.Text = Get-Val "phone"
        $global:createUserDescriptionTextBox.Text = Get-Val "description"
        
        # Organization mapping
        $company = Get-Val "company"
        if ($company) {
            switch -Regex ($company.ToLower()) {
                "hra"   { $global:createUserOrgComboBox.SelectedItem = $global:createUserOrgComboBox.Items | Where-Object { $_.Content -eq "HRA" } }
                "dss"   { $global:createUserOrgComboBox.SelectedItem = $global:createUserOrgComboBox.Items | Where-Object { $_.Content -eq "DSS" } }
                "azure" { $global:createUserOrgComboBox.SelectedItem = $global:createUserOrgComboBox.Items | Where-Object { $_.Content -eq "Azure" } }
            }
        }
        
        # OU mapping from CSV (if present)
        $csvOU = Get-Val "ou"
        if ($csvOU) {
            switch -Regex ($csvOU.ToLower()) {
                "accis" { 
                    $selectedItem = $global:createUserOUComboBox.Items | Where-Object { $_.Content -match "ACCIS" } | Select-Object -First 1
                    if ($selectedItem) { $global:createUserOUComboBox.SelectedItem = $selectedItem }
                }
                "mcs"   { 
                    $selectedItem = $global:createUserOUComboBox.Items | Where-Object { $_.Content -match "MCS" } | Select-Object -First 1
                    if ($selectedItem) { $global:createUserOUComboBox.SelectedItem = $selectedItem }
                }
                "pactweb" { 
                    $selectedItem = $global:createUserOUComboBox.Items | Where-Object { $_.Content -match "PACTWEB" } | Select-Object -First 1
                    if ($selectedItem) { $global:createUserOUComboBox.SelectedItem = $selectedItem }
                }
                "vsp-web|stars" { 
                    $selectedItem = $global:createUserOUComboBox.Items | Where-Object { $_.Content -match "VSP-Web.*STARS" } | Select-Object -First 1
                    if ($selectedItem) { $global:createUserOUComboBox.SelectedItem = $selectedItem }
                }
            }
        }
        
        $global:outputTextBox.AppendText("CSV loaded successfully. Review before creating user.`n")
    }
    catch {
        $global:outputTextBox.AppendText("CSV Load Failed: $_`n")
    }
}

# Event handler for Create User Submit button
$global:createUserSubmitButton.Add_Click({
    $global:outputTextBox.Clear()
    $global:outputTextBox.AppendText("Starting user creation process...`n")
    
    # Force UI update
    $global:window.Dispatcher.Invoke([action]{
        try {
            # Get values from form
            $firstName = $global:createUserFirstNameTextBox.Text.Trim()
            $lastName = $global:createUserLastNameTextBox.Text.Trim()
            $title = $global:createUserTitleTextBox.Text.Trim()
            $department = $global:createUserDepartmentTextBox.Text.Trim()
            $phone = $global:createUserPhoneTextBox.Text.Trim()
            $email = $global:createUserEmailTextBox.Text.Trim()
            $description = $global:createUserDescriptionTextBox.Text.Trim()
            $groups = $global:createUserGroupsTextBox.Text.Trim()
            $manualLanID = $global:createUserLanIDTextBox.Text.Trim()
            $ouSelection = if ($global:createUserOUComboBox.SelectedItem) { $global:createUserOUComboBox.SelectedItem.Content } else { "(Default)" }
            $org = if ($global:createUserOrgComboBox.SelectedItem) { $global:createUserOrgComboBox.SelectedItem.Content } else { "" }
            $type = if ($global:createUserTypeComboBox.SelectedItem) { $global:createUserTypeComboBox.SelectedItem.Content } else { "" }
            $expiry = if ($global:createUserExpiryComboBox.SelectedItem) { [int]$global:createUserExpiryComboBox.SelectedItem.Content } else { 6 }
            
            # Validate required fields
            if ([string]::IsNullOrEmpty($firstName) -or [string]::IsNullOrEmpty($lastName)) {
                $global:outputTextBox.AppendText("Error: First Name and Last Name are required.`n")
                return
            }
            
            if ([string]::IsNullOrEmpty($org)) {
                $global:outputTextBox.AppendText("Error: Organization is required.`n")
                return
            }
            
            if ([string]::IsNullOrEmpty($type)) {
                $global:outputTextBox.AppendText("Error: Account Type is required.`n")
                return
            }
            
            # Generate LAN ID based on rules
            $global:outputTextBox.AppendText("Generating LAN ID...`n")
            $lanID = Generate-LanID -LastName $lastName -ManualLanID $manualLanID
            $global:outputTextBox.AppendText("Generated LAN ID: $lanID`n")
            
            # Check for duplicates BEFORE creating
            $global:outputTextBox.AppendText("Checking for duplicate users...`n")
            $duplicates = Check-ForDuplicateUsers -FirstName $firstName -LastName $lastName -ProposedLanID $lanID
            
            if ($duplicates.Count -gt 0) {
                $global:outputTextBox.AppendText("=" * 60 + "`n")
                $global:outputTextBox.AppendText("DUPLICATE USERS FOUND!`n")
                $global:outputTextBox.AppendText("=" * 60 + "`n")
                
                foreach ($dup in $duplicates) {
                    $status = if ($dup.Enabled) { "ENABLED" } else { "DISABLED" }
                    $global:outputTextBox.AppendText("Type: $($dup.Type)`n")
                    $global:outputTextBox.AppendText("Name: $($dup.DisplayName)`n")
                    $global:outputTextBox.AppendText("Username: $($dup.SamAccountName)`n")
                    $global:outputTextBox.AppendText("UPN: $($dup.UserPrincipalName)`n")
                    $global:outputTextBox.AppendText("Status: $status`n")
                    $global:outputTextBox.AppendText("-" * 40 + "`n")
                }
                
                $global:outputTextBox.AppendText("=" * 60 + "`n")
                $global:outputTextBox.AppendText("User creation CANCELLED due to duplicates.`n")
                $global:outputTextBox.AppendText("Please review the duplicates above.`n")
                return
            } else {
                $global:outputTextBox.AppendText("No duplicates found. Proceeding with creation...`n")
            }
            
            # Determine account type code for EmployeeType attribute
            $accountCode = switch ($type) {
                "Employee (E)" { "E" }
                "Temp (T)" { "T" }
                "B2B (B)" { "B" }
                default { "E" }
            }
            
            # Determine OU based on selection (using pattern matching)
            switch ($ouSelection) {
                {$_ -match "ACCIS"} {
                    $targetOU = "OU=ACCIS-AzureAD,OU=DMZ-AzureAD,OU=DMZ,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
                }
                {$_ -match "MCS"} {
                    $targetOU = "OU=MCS-AzureAD,OU=DMZ-AzureAD,OU=DMZ,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
                }
                {$_ -match "PACTWEB"} {
                    $targetOU = "OU=PACTWEB-AzureAD,OU=DMZ-AzureAD,OU=DMZ,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
                }
                {$_ -match "VSP-Web.*STARS"} {
                    $targetOU = "OU=VSP-Web_STARS_SEAMS-AzureAD,OU=DMZ-AzureAD,OU=DMZ,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
                }
                default {
                    $targetOU = $DefaultOU
                }
            }
            
            # Create display name
            $displayName = "$firstName $lastName"
            
            # Determine UPN suffix based on organization
            switch ($org) {
                "Azure" {
                    $upnSuffix = "@nychra.onmicrosoft.com"
                    $changePasswordAtLogon = $false
                }
                default { # HRA and DSS
                    $upnSuffix = "@windows.nyc.hra.nycnet"
                    $changePasswordAtLogon = $true
                }
            }
            
            # Generate UserPrincipalName
            $upn = "$lanID$upnSuffix"
            
            # Calculate account expiration
            $accountExpiration = $null
            if ($type -eq "Temp (T)" -and $expiry -gt 0) {
                $accountExpiration = (Get-Date).AddMonths($expiry)
                $global:outputTextBox.AppendText("Account will expire on: $accountExpiration`n")
            }
            
            # Create user parameters
            $userParams = @{
                SamAccountName = $lanID
                Name = $displayName
                GivenName = $firstName
                Surname = $lastName
                DisplayName = $displayName
                UserPrincipalName = $upn
                EmailAddress = $email
                Title = $title
                Department = $department
                Description = $description
                OfficePhone = $phone
                AccountPassword = (ConvertTo-SecureString $DefaultPassword -AsPlainText -Force)
                Enabled = $true
                ChangePasswordAtLogon = $changePasswordAtLogon
                Path = $targetOU
            }
            
            if ($accountExpiration) {
                $userParams.AccountExpirationDate = $accountExpiration
            }
            
            # Create the user
            $global:outputTextBox.AppendText("`nCreating user '$lanID' in OU: $targetOU`n")
            $global:outputTextBox.AppendText("UPN: $upn`n")
            $global:outputTextBox.AppendText("Password change at logon: $changePasswordAtLogon`n")
            
            New-ADUser @userParams -ErrorAction Stop
            $global:outputTextBox.AppendText("User '$lanID' created successfully.`n")
            
            # Add to Internet group
            try {
                Add-ADGroupMember -Identity $InternetGroup -Members $lanID -ErrorAction Stop
                $global:outputTextBox.AppendText("Added to Internet group: $InternetGroup`n")
            } catch {
                $global:outputTextBox.AppendText("Warning: Could not add to Internet group: $_`n")
            }
            
            # Add to extra groups if specified
            if (-not [string]::IsNullOrEmpty($groups)) {
                $groupArray = $groups -split "," | ForEach-Object { $_.Trim() }
                foreach ($group in $groupArray) {
                    if (-not [string]::IsNullOrEmpty($group)) {
                        try {
                            Add-ADGroupMember -Identity $group -Members $lanID -ErrorAction Stop
                            $global:outputTextBox.AppendText("Added to group: $group`n")
                        } catch {
                            $global:outputTextBox.AppendText("Warning: Could not add to group '$group': $_`n")
                        }
                    }
                }
            }
            
            # Set EmployeeType attribute and organization extension attribute
            $userAttributes = @{
                EmployeeType = $accountCode
            }
            
            if ($org) {
                $userAttributes.extensionAttribute1 = $org
            }
            
            Set-ADUser -Identity $lanID -Replace $userAttributes
            $global:outputTextBox.AppendText("Set Employee Type: $accountCode`n")
            if ($org) {
                $global:outputTextBox.AppendText("Set Organization: $org`n")
            }
            
            $global:outputTextBox.AppendText("`n" + "=" * 40 + "`n")
            $global:outputTextBox.AppendText("USER CREATION COMPLETE`n")
            $global:outputTextBox.AppendText("Username: $lanID`n")
            $global:outputTextBox.AppendText("Password: $DefaultPassword`n")
            $global:outputTextBox.AppendText("UPN: $upn`n")
            $global:outputTextBox.AppendText("OU: $targetOU`n")
            $global:outputTextBox.AppendText("Organization: $org`n")
            $global:outputTextBox.AppendText("Employee Type: $accountCode`n")
            $global:outputTextBox.AppendText("Password change at logon: $changePasswordAtLogon`n")
            $global:outputTextBox.AppendText("=" * 40 + "`n")
            
            # Clear form on success
            $global:createUserFirstNameTextBox.Clear()
            $global:createUserLastNameTextBox.Clear()
            $global:createUserTitleTextBox.Clear()
            $global:createUserDepartmentTextBox.Clear()
            $global:createUserPhoneTextBox.Clear()
            $global:createUserEmailTextBox.Clear()
            $global:createUserDescriptionTextBox.Clear()
            $global:createUserGroupsTextBox.Clear()
            $global:createUserLanIDTextBox.Clear()
            $global:createUserOrgComboBox.SelectedIndex = -1
            $global:createUserTypeComboBox.SelectedIndex = -1
            $global:createUserOUComboBox.SelectedIndex = 0  # Reset to "(Default)"
            $global:createUserExpiryLabel.Visibility = "Collapsed"
            $global:createUserExpiryComboBox.Visibility = "Collapsed"
            $global:createUserExpiryComboBox.SelectedIndex = 0
            
        } catch {
            $global:outputTextBox.AppendText("Error creating user: $_`n")
        }
    }, "Normal")
})

# Event handler for Load CSV button
$global:createUserLoadCSVButton.Add_Click({
    $global:outputTextBox.Clear()
    $global:outputTextBox.AppendText("CSV Import started...`n")
    Import-UserFromCSV
})

# Event handler for Clear Form button
$global:createUserClearButton.Add_Click({
    $global:createUserFirstNameTextBox.Clear()
    $global:createUserLastNameTextBox.Clear()
    $global:createUserTitleTextBox.Clear()
    $global:createUserDepartmentTextBox.Clear()
    $global:createUserPhoneTextBox.Clear()
    $global:createUserEmailTextBox.Clear()
    $global:createUserDescriptionTextBox.Clear()
    $global:createUserGroupsTextBox.Clear()
    $global:createUserLanIDTextBox.Clear()
    $global:createUserOrgComboBox.SelectedIndex = -1
    $global:createUserTypeComboBox.SelectedIndex = -1
    $global:createUserOUComboBox.SelectedIndex = 0  # Reset to "(Default)"
    $global:createUserExpiryLabel.Visibility = "Collapsed"
    $global:createUserExpiryComboBox.Visibility = "Collapsed"
    $global:createUserExpiryComboBox.SelectedIndex = 0
    $global:outputTextBox.AppendText("Form cleared.`n")
})

# Event handler for Account Type change (show/hide expiry)
$global:createUserTypeComboBox.Add_SelectionChanged({
    if ($global:createUserTypeComboBox.SelectedItem) {
        $isTemp = $global:createUserTypeComboBox.SelectedItem.Content -match "Temp"
        $global:createUserExpiryLabel.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
        $global:createUserExpiryComboBox.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
    }
})

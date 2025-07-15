<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName System.Drawing

# Excel COM Object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Get the directory of the currently running script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the relative path to the XAML file
$xamlFilePath = Join-Path -Path $scriptDir -ChildPath "MainWindow.xaml"

# Load the XAML
[xml]$xaml = Get-Content -Path $xamlFilePath
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$remoteAccessButton = $window.FindName("RemoteAccessButton")
$lanExtensionButton = $window.FindName("LanExtensionButton")
$groupManagementButton = $window.FindName("GroupManagementButton")
$MFAUpdateButton = $window.FindName("MFAUpdateButton")
$remoteAccessPanel = $window.FindName("RemoteAccessPanel")
$lanExtensionPanel = $window.FindName("LanExtensionPanel")
$groupManagementPanel = $window.FindName("GroupManagementPanel")
$MFAUpdatePanel = $window.FindName("MFAUpdatePanel")
$outputTextBox = $window.FindName("OutputTextBox")
$remoteAccessSubmitButton = $window.FindName("RemoteAccessSubmitButton")
$lanExtensionSubmitButton = $window.FindName("LanExtensionSubmitButton")
$groupManagementSubmitButton = $window.FindName("GroupManagementSubmitButton")
$MFAUpdateSubmitButton = $window.FindName("MFAUpdateSubmitButton")

# Add new controls for Group Extraction
$groupExtractionButton = $window.FindName("GroupExtractionButton")
$groupExtractionPanel = $window.FindName("GroupExtractionPanel")
$groupNameTextBox = $window.FindName("GroupNameTextBox")
$extractButton = $window.FindName("ExtractButton")
$downloadButton = $window.FindName("DownloadButton")
$progressBar = $window.FindName("ProgressBar")
$statusLabel = $window.FindName("StatusLabel")

# Initialize variables for extracted data
$extractedData = $null
$exportFilePath = $null

# ===== EVENT HANDLERS =====

# Panel visibility handlers
$remoteAccessButton.Add_Click({
    $remoteAccessPanel.Visibility = "Visible"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

$lanExtensionButton.Add_Click({
    $lanExtensionPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

$groupManagementButton.Add_Click({
    $groupManagementPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

$MFAUpdateButton.Add_Click({
    $MFAUpdatePanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

$groupExtractionButton.Add_Click({
    $groupExtractionPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})




# Event handler for Remote Access Submit button
$remoteAccessSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $computerName = $window.FindName("ComputerNameTextBox").Text
    $userLANID = $window.FindName("UserLANIDTextBox").Text
    $groupSelection = $window.FindName("GroupSelectionComboBox").SelectedItem.Content

    if ([string]::IsNullOrEmpty($computerName) -or [string]::IsNullOrEmpty($userLANID) -or [string]::IsNullOrEmpty($groupSelection)) {
        $outputTextBox.AppendText("Error: Computer Name, User LAN ID, and Group Selection are required.`n")
        return
    }

    # Map the selection to the actual group name
    $groupToAdd = switch ($groupSelection) {
        "1" { "LUW-HRAPersonalRDPUsers" }
        "2" { "HRARDPUsers2" }
        "3" { "HRARDPUsers3" }
        default { "HRARDPUsers3" } # Default to 3 if something unexpected comes through
    }

    try {
        # --- ADD USER TO SELECTED GROUP (IF NOT ALREADY A MEMBER) ---
        $user = Get-ADUser -Identity $userLANID -Properties memberof -ErrorAction Stop

        # Check if user is already in the selected group
        $groupDN = (Get-ADGroup -Identity $groupToAdd -ErrorAction Stop).DistinguishedName
        $isMember = $user.memberof -contains $groupDN

        if (-not $isMember) {
            Add-ADGroupMember -Identity $groupToAdd -Members $userLANID -ErrorAction Stop
            $outputTextBox.AppendText("SUCCESS: User added to $groupToAdd.`n")
        } else {
            $outputTextBox.AppendText("INFO: User is already in $groupToAdd.`n")
        }

        # --- REST OF YOUR EXISTING SCRIPT CONTINUES HERE ---
        # (The code for removing from unwanted groups, computer checks, etc.)
        
        # --- REMOVE USER FROM UNWANTED GROUPS (IF THEY EXIST) ---
        $groupsToRemove = @("GS-MFA-IPADusers", "GS-MFA-MACusers", "GS-MFA-NewRadiusAuthentication")
        $removedGroups = @()

        foreach ($group in $groupsToRemove) {
            try {
                $groupObj = Get-ADGroup -Identity $group -ErrorAction Stop
                $groupDN = $groupObj.DistinguishedName

                # Check if user is a member before removing
                if ($user.memberof -contains $groupDN) {
                    Remove-ADGroupMember -Identity $group -Members $userLANID -Confirm:$false -ErrorAction Stop
                    $removedGroups += $group
                    $outputTextBox.AppendText("SUCCESS: User removed from $group.`n")
                }
            } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $outputTextBox.AppendText("INFO: Group $group does not exist (skipped).`n")
            } catch {
                $outputTextBox.AppendText("ERROR: Failed to process $group - $($_.Exception.Message)`n")
            }
        }

        # Summary
        if ($removedGroups.Count -gt 0) {
            $outputTextBox.AppendText("SUMMARY: User was removed from: $($removedGroups -join ', ').`n")
        } else {
            $outputTextBox.AppendText("SUMMARY: No groups were removed (user was not a member or groups didn't exist).`n")
        }

        # Check if the computer exists and ping to check if it's online
        $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop
        if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
            $outputTextBox.AppendText("The computer $computerName is online.`n")

            # Remove current users from Remote Desktop Users group and add the specified user
            Invoke-Command -ComputerName $computerName -ScriptBlock {
                $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                $members = @($group.PSBase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $_, $null) }
                foreach ($member in $members) {
                    $group.Remove($member)
                }
                $group.Add("WinNT://$using:userLANID,user")
            }
            $outputTextBox.AppendText("The user $userLANID has been added to the Remote Desktop Users group on $computerName.`n")

            # Update extensionAttribute2 for the user
            $user = Get-ADUser -Identity $userLANID -Properties extensionAttribute2
            $user.extensionAttribute2 = "$computerName.windows.nyc.hra.nycnet"
            Set-ADUser -Identity $userLANID -Replace @{extensionAttribute2 = $user.extensionAttribute2}
            $outputTextBox.AppendText("Close Out Ticket: User has been given access to PC.`n")
        } else {
            $outputTextBox.AppendText("The computer $computerName is found but offline.`n")

            # Remove extensionAttribute2 for the user if the computer is offline
            Set-ADUser -Identity $userLANID -Clear extensionAttribute2
            $outputTextBox.AppendText("Extension attribute has been removed.`n")
            $outputTextBox.AppendText("Close Out Ticket: PC seems to be offline. Please check the connection of the PC.`n")
        }
    } catch {
        $outputTextBox.AppendText("The computer $computerName does not exist.`n")

        # Remove extensionAttribute2 for the user if the computer does not exist
        Set-ADUser -Identity $userLANID -Clear extensionAttribute2
        $outputTextBox.AppendText("Extension attribute has been removed.`n")
        $outputTextBox.AppendText("Close Out Ticket: PC was not able to be located. Please verify the PC name.`n")
    }
})

# Event handler for LAN Extension Submit button
$lanExtensionSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $userLanId = $window.FindName("LanExtensionUserLANIDTextBox").Text
    $extendDate = $window.FindName("LanExtensionDateTextBox").Text
    $ticketNumber = $window.FindName("LanExtensionTicketNumberTextBox").Text
    $initials = $window.FindName("LanExtensionInitialsTextBox").Text

    if ([string]::IsNullOrEmpty($userLanId) -or [string]::IsNullOrEmpty($extendDate) -or 
        [string]::IsNullOrEmpty($ticketNumber) -or [string]::IsNullOrEmpty($initials)) {
        $outputTextBox.AppendText("Error: All fields are required.`n")
        return
    }

    try {
        # Import the Active Directory module
        Import-Module ActiveDirectory

        # Find the user in Active Directory using the correct LAN ID variable
        $user = Get-ADUser -Filter {SamAccountName -eq $userLanId} -Properties AccountExpirationDate, UserPrincipalName, Enabled, Description

        # Check if the user was found
        if ($null -eq $user) {
            $outputTextBox.AppendText("User not found.`n")
            return
        }


        # Check if the user's account is disabled and enable it if necessary
        if ($user.Enabled -eq $false) {
            Set-ADUser -Identity $user -Enabled $true
            $outputTextBox.AppendText("User account was disabled and has now been enabled.`n")
        } else {
            $outputTextBox.AppendText("User account is enabled.`n")
        }

        # Check if the "Account never expires" box is ticked
        if ($user.AccountExpirationDate -eq $null) {
            # Untick the "Account never expires" box by setting an expiration date
            Set-ADUser -Identity $user -AccountExpirationDate (Get-Date -Date "12/31/9999") # Temporary date to untick the box
        }

    

        # Calculate the new extension date by adding one day to the given date
        $newExtendDate = (Get-Date -Date $extendDate).AddDays(1)

        if (-not $isDHS) {
            # Update the user's account expiration date
            Set-ADUser -Identity $user -AccountExpirationDate $newExtendDate

            # Move the user to the new OU
            $newOU = "OU=Temps (Replacing 15 MTC Temps OU),OU=470 Vanderbilt,OU=People,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $newOU
            $outputTextBox.AppendText("User's account has been extended to $extendDate and moved to 470 Vanderbilt Temps OU.`n")
        } else {
            $outputTextBox.AppendText("User is DHS. Account not extended and location not moved.`n")
        }

        # Update the user's description field by prepending the new description
        $newDescription = "Extended as per $ticketNumber $initials | "
        $existingDescription = $user.Description
        Set-ADUser -Identity $user -Description "$newDescription$existingDescription"

        $outputTextBox.AppendText("Ticket Close Test: User's account has been extended to $extendDate.`n")

    } catch {
        $outputTextBox.AppendText("Error: $_`n")
    } finally {
        # Clear the fields after submission
        $window.FindName("LanExtensionUserLANIDTextBox").Text = ""
        $window.FindName("LanExtensionDateTextBox").Text = ""
        $window.FindName("LanExtensionTicketNumberTextBox").Text = ""
        # Do not clear the Initials field
    }
})

# Event handler for Group Management Submit button
$groupManagementSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $userLanIDs = $window.FindName("GroupManagementUserLANIDTextBox").Text
    $groupNames = $window.FindName("GroupManagementGroupNamesTextBox").Text
    $action = $window.FindName("GroupManagementActionComboBox").SelectedItem.Content

    if ([string]::IsNullOrEmpty($userLanIDs) -or [string]::IsNullOrEmpty($groupNames)) {
        $outputTextBox.AppendText("Error: Both User LAN ID(s) and Group Name(s) are required.`n")
        return
    }

    $userLanIDsArray = $userLanIDs -split ","
    $groupNamesArray = $groupNames -split ","

    foreach ($userLanID in $userLanIDsArray) {
        if ($action -eq "Add") {
            Add-UserToGroups -UserLanID $userLanID.Trim() -GroupNames $groupNamesArray
        } elseif ($action -eq "Remove") {
            Remove-UserFromGroups -UserLanID $userLanID.Trim() -GroupNames $groupNamesArray
        } else {
            $outputTextBox.AppendText("Invalid action selected.`n")
        }
    }
})

# Event handler for MFA Update Submit button
$MFAUpdateSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $userEmail = $window.FindName("MFAUpdateUserEmailTextBox").Text
    $phoneNumber = $window.FindName("MFAUpdatePhoneNumberTextBox").Text
    $methodType = $window.FindName("MFAUpdateMethodTypeComboBox").SelectedItem.Content

    if ([string]::IsNullOrEmpty($userEmail) -or [string]::IsNullOrEmpty($phoneNumber)) {
        $outputTextBox.AppendText("Error: Both User Email and Phone Number are required.`n")
        return
    }

    try {
        # Connect to Microsoft Graph
        Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All"

        # Call the function to update the phone number
        Update-AuthenticationPhoneNumber -UserEmail $userEmail -PhoneNumber $phoneNumber -MethodType $methodType
        $outputTextBox.AppendText("Close out ticket: User phone number has been updated.`n")
    } catch {
        $outputTextBox.AppendText("An error occurred: $_`n")
        $outputTextBox.AppendText("Please activate your roles and try again.`n")
    }
})

# Group Extraction handlers
$extractButton.Add_Click({
    $groupName = $groupNameTextBox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($groupName)) {
        $outputTextBox.AppendText("Error: Group name is required.`n")
        return
    }
    
    try {
        $outputTextBox.AppendText("Starting group extraction...`n")
        $progressBar.Value = 10
        $statusLabel.Content = "Connecting to Active Directory..."
        
        # Check if group exists
        $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction Stop
        if (-not $group) {
            $outputTextBox.AppendText("Error: Group '$groupName' not found in Active Directory.`n")
            $progressBar.Value = 0
            $statusLabel.Content = "Ready"
            return
        }
        
        $progressBar.Value = 20
        $statusLabel.Content = "Retrieving group members..."
        $outputTextBox.AppendText("Group found. Retrieving members...`n")
        
        # Get all members of the group
        $members = Get-ADGroupMember -Identity $groupName -Recursive | 
                   Where-Object { $_.objectClass -eq 'user' } | 
                   Get-ADUser -Properties *
        
        $progressBar.Value = 40
        $statusLabel.Content = "Processing user data..."
        $outputTextBox.AppendText("Found $($members.Count) users. Processing data...`n")
        
        # Extract user information
        $script:extractedData = @()
        $userCount = $members.Count
        $currentUser = 0
        
        foreach ($user in $members) {
            $currentUser++
            $progress = 40 + ([math]::Round(($currentUser / $userCount) * 50))
            $progressBar.Value = $progress
            $statusLabel.Content = "Processing user $currentUser of $userCount..."
            
            $userInfo = [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                UserPrincipalName = $user.UserPrincipalName
                Email = $user.EmailAddress
                Title = $user.Title
                Department = $user.Department
                Company = $user.Company
                Office = $user.Office
                StreetAddress = $user.StreetAddress
                City = $user.City
                State = $user.State
                PostalCode = $user.PostalCode
                Country = $user.Country
                Telephone = $user.telephoneNumber
                Mobile = $user.mobile
                EmployeeID = $user.EmployeeID
                EmployeeType = $user.EmployeeType
                Enabled = $user.Enabled
                LastLogonDate = $user.LastLogonDate
                PasswordLastSet = $user.PasswordLastSet
                AccountExpirationDate = $user.AccountExpirationDate
                WhenCreated = $user.WhenCreated
                DistinguishedName = $user.DistinguishedName
            }
            
            $script:extractedData += $userInfo
        }
        
        $progressBar.Value = 100
        $statusLabel.Content = "Extraction complete"
        $outputTextBox.AppendText("Successfully extracted data for $($script:extractedData.Count) users.`n")
        $downloadButton.IsEnabled = $true
        
    } catch {
        $outputTextBox.AppendText("Error during extraction: $_`n")
        $progressBar.Value = 0
        $statusLabel.Content = "Error occurred"
    }
})

$downloadButton.Add_Click({
    if (-not $script:extractedData -or $script:extractedData.Count -eq 0) {
        $outputTextBox.AppendText("Error: No data to export. Please extract data first.`n")
        return
    }
    
    try {
        $progressBar.Value = 0
        $statusLabel.Content = "Preparing export..."
        $outputTextBox.AppendText("Preparing Excel export...`n")
        
        # Create a SaveFileDialog
        $saveFileDialog = New-Object Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $saveFileDialog.FileName = "$($groupNameTextBox.Text)_UserExport_$(Get-Date -Format 'yyyyMMdd').xlsx"
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            $script:exportFilePath = $saveFileDialog.FileName
            $outputTextBox.AppendText("Exporting to: $script:exportFilePath`n")
            
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            
            # Create Excel objects
            $progressBar.Value = 10
            $statusLabel.Content = "Creating Excel file..."
            
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "User Data"
            
            # Add headers
            $progressBar.Value = 20
            $statusLabel.Content = "Adding headers..."
            $column = 1
            $script:extractedData[0].PSObject.Properties.Name | ForEach-Object {
                $worksheet.Cells.Item(1, $column) = $_
                $worksheet.Cells.Item(1, $column).Font.Bold = $true
                $worksheet.Cells.Item(1, $column).Interior.ColorIndex = 15
                $column++
            }
            
            # Add data
            $progressBar.Value = 30
            $statusLabel.Content = "Adding data..."
            $row = 2
            $totalRows = $script:extractedData.Count
            $currentRow = 0
            
            foreach ($user in $script:extractedData) {
                $currentRow++
                $progress = 30 + ([math]::Round(($currentRow / $totalRows) * 60))
                $progressBar.Value = $progress
                $statusLabel.Content = "Exporting row $currentRow of $totalRows..."
                
                $column = 1
                $user.PSObject.Properties.Value | ForEach-Object {
                    $worksheet.Cells.Item($row, $column) = $_
                    $column++
                }
                $row++
            }
            
            # Auto-fit columns
            $progressBar.Value = 95
            $statusLabel.Content = "Formatting document..."
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
            
            # Add freeze panes and formatting
            $worksheet.Activate()
            $worksheet.Application.ActiveWindow.SplitRow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true
            
            # Save and close
            $progressBar.Value = 98
            $statusLabel.Content = "Saving file..."
            $workbook.SaveAs($script:exportFilePath)
            $workbook.Close($false)
            
            $progressBar.Value = 100
            $statusLabel.Content = "Export complete"
            $outputTextBox.AppendText("Successfully exported data to Excel file.`n")
            
            # Offer to open the file
            $result = [System.Windows.MessageBox]::Show("Export completed successfully. Would you like to open the file now?", "Export Complete", "YesNo", "Question")
            if ($result -eq "Yes") {
                Start-Process $script:exportFilePath
            }
        }
    } catch {
        $outputTextBox.AppendText("Error during export: $_`n")
        $progressBar.Value = 0
        $statusLabel.Content = "Error occurred"
    } finally {
        if ($workbook) { $workbook.Close($false) }
        $progressBar.Value = 0
        $statusLabel.Content = "Ready"
    }
})

# ===== SHOW WINDOW =====
$window.ShowDialog() | Out-Null

# Clean up Excel
try {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} catch {
    Write-Warning "Error cleaning up Excel COM objects: $_"
}

# Function to add a user to one or multiple groups
function Add-UserToGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLanID,
        [Parameter(Mandatory = $true)]
        [string[]]$GroupNames
    )

    $User = Get-ADUser -Filter "SamAccountName -eq '$UserLanID'" -Properties SamAccountName, MemberOf
    if ($null -eq $User) {
        $outputTextBox.AppendText("User with LAN ID '$UserLanID' not found in Active Directory.`n")
        return
    } else {
        $outputTextBox.AppendText("User '$UserLanID' found.`n")
    }

    foreach ($GroupName in $GroupNames) {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -Properties Name
        if ($null -eq $Group) {
            $outputTextBox.AppendText("Group '$GroupName' does not exist in Active Directory.`n")
            continue
        } else {
            $outputTextBox.AppendText("Group '$GroupName' found.`n")
        }

        $IsMember = $User.MemberOf -contains ("CN=$GroupName,OU=Groups,DC=YourDomain,DC=com")
        if ($IsMember) {
            $outputTextBox.AppendText("User '$UserLanID' is already a member of the group '$GroupName'.`n")
        } else {
            Add-ADGroupMember -Identity $GroupName -Members $User.DistinguishedName
            $outputTextBox.AppendText("User '$UserLanID' has been added to the group '$GroupName'.`n")
        }
    }
}

# Function to remove a user from one or multiple groups
function Remove-UserFromGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLanID,
        [Parameter(Mandatory = $true)]
        [string[]]$GroupNames
    )

    $User = Get-ADUser -Filter "SamAccountName -eq '$UserLanID'" -Properties SamAccountName
    if ($null -eq $User) {
        $outputTextBox.AppendText("User with LAN ID '$UserLanID' not found in Active Directory.`n")
        return
    } else {
        $outputTextBox.AppendText("User '$UserLanID' found.`n")
    }

    foreach ($GroupName in $GroupNames) {
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -Properties Name
        if ($null -eq $Group) {
            $outputTextBox.AppendText("Group '$GroupName' does not exist in Active Directory.`n")
            continue
        } else {
            $outputTextBox.AppendText("Group '$GroupName' found.`n")
        }

        $IsMember = Get-ADGroupMember -Identity $GroupName | Where-Object { $_.SamAccountName -eq $UserLanID }
        if ($IsMember) {
            Remove-ADGroupMember -Identity $GroupName -Members $User.DistinguishedName -Confirm:$false
            $outputTextBox.AppendText("User '$UserLanID' has been removed from the group '$GroupName'.`n")
        } else {
            $outputTextBox.AppendText("User '$UserLanID' is no longer a member of the group '$GroupName'.`n")
        }
    }
}

# Function to format phone number
function Format-PhoneNumber {
    param (
        [string]$PhoneNumber
    )
    # Remove non-numeric characters except for the plus sign
    $formattedNumber = $PhoneNumber -replace '[^\d\+]', ''
    # Ensure the phone number starts with "+1"
    if (-not $formattedNumber.StartsWith("+1")) {
        $formattedNumber = "+1" + $formattedNumber.TrimStart('+', '1')
    }
    return $formattedNumber
}

# Function to update phone numbers in the authentication methods section
function Update-AuthenticationPhoneNumber {
    param (
        [string]$UserEmail,
        [string]$PhoneNumber,
        [string]$MethodType
    )
    
    try {
        # Get the user using the provided email
        $user = Get-MgUser -Filter "mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'"
        if (-not $user) {
            $outputTextBox.AppendText("User not found!`n")
            return
        }
        
        # Format number
        $formattedPhoneNumber = Format-PhoneNumber -PhoneNumber $PhoneNumber

        # phone method type based on user input
        switch ($MethodType) {
            "Mobile" { $MethodType = "mobile" }
            "Alternate Mobile" { $MethodType = "alternateMobile" }
            "Office" { $MethodType = "office" }
            default { $MethodType = $MethodType.ToLower() }
        }

        # Perform the action
        switch ($MethodType) {
            "mobile" {
                # Check if the user already has a primary mobile phone number
                $existingPhoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id | Where-Object { $_.PhoneType -eq "mobile" }
                if ($existingPhoneMethod) {
                    # Update the existing mobile phone number
                    Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber
                    $outputTextBox.AppendText("Close out ticket: User phone number has been changed.`n")
                } else {
                    # Add a new mobile phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "mobile"
                    $outputTextBox.AppendText("Close out ticket: User phone number has been added.`n")
                }
            }
            "alternateMobile" {
                # Check if the user already has an alternate phone number
                $existingPhoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id | Where-Object { $_.PhoneType -eq "alternateMobile" }
                if ($existingPhoneMethod) {
                    # Update the existing alternate phone number
                    Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber
                    $outputTextBox.AppendText("Close out ticket: User phone number has been changed.`n")
                } else {
                    # Add a new alternate phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "alternateMobile"
                    $outputTextBox.AppendText("Close out ticket: User phone number has been added.`n")
                }
            }
            "office" {
                # Check if the user already has an office phone number
                $existingPhoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id | Where-Object { $_.PhoneType -eq "office" }
                if ($existingPhoneMethod) {
                    # Update the existing office phone number
                    Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber
                    $outputTextBox.AppendText("Close out ticket: User phone number has been changed.`n")
                } else {
                    # Add a new office phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "office"
                    $outputTextBox.AppendText("Close out ticket: User phone number has been added.`n")
                }
            }
            default {
                $outputTextBox.AppendText("Invalid phone method type specified!`n")
            }
        }
    } catch {
        $outputTextBox.AppendText("An error occurred: $_`n")
        $outputTextBox.AppendText("Please activate your roles and try again.`n")
    }
}

# Show the window
$window.ShowDialog() | Out-Null

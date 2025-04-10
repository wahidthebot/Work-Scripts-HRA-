<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>
 
Add-Type -AssemblyName PresentationFramework

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

# Event handler for Remote Access button
$remoteAccessButton.Add_Click({
    $remoteAccessPanel.Visibility = "Visible"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

# Event handler for LAN Extension button
$lanExtensionButton.Add_Click({
    $lanExtensionPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

# Event handler for Group Management button
$groupManagementButton.Add_Click({
    $groupManagementPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

# Event handler for MFA Update button
$MFAUpdateButton.Add_Click({
    $MFAUpdatePanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
})

# Event handler for Remote Access Submit button
$remoteAccessSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $computerName = $window.FindName("ComputerNameTextBox").Text
    $userLANID = $window.FindName("UserLANIDTextBox").Text

    if ([string]::IsNullOrEmpty($computerName) -or [string]::IsNullOrEmpty($userLANID)) {
        $outputTextBox.AppendText("Error: Both Computer Name and User LAN ID are required.`n")
        return
    }

 try {
    # --- ADD USER TO HRARDPUsers3 (IF NOT ALREADY A MEMBER) ---
    $groupToAdd = "HRARDPUsers3"
    $user = Get-ADUser -Identity $userLANID -Properties memberof -ErrorAction Stop

    # Check if user is already in HRARDPUsers3
    $groupDN = (Get-ADGroup -Identity $groupToAdd -ErrorAction Stop).DistinguishedName
    $isMember = $user.memberof -contains $groupDN

    if (-not $isMember) {
        Add-ADGroupMember -Identity $groupToAdd -Members $userLANID -ErrorAction Stop
        $outputTextBox.AppendText("SUCCESS: User added to $groupToAdd.`n")
    } else {
        $outputTextBox.AppendText("INFO: User is already in $groupToAdd.`n")
    }

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

    # Output results
    if ($notMembers.Count -gt 0) {
        $outputTextBox.AppendText("The user was added to the following groups: $($notMembers -join ', ')`n")
    } else {
        $outputTextBox.AppendText("The user is already a member of all required groups.`n")
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

        # Check if the user is from DHS
        $isDHS = $user.UserPrincipalName -like "*@dhs.nyc.gov"

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

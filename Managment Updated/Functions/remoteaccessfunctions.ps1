# ===== GLOBAL VARIABLES TO STORE COMBOBOX VALUES =====
$global:SelectedAccountType = "Default"
# PC Farm Group variable is no longer needed since all users go to the same group

# ===== COMBOBOX EVENT HANDLERS =====
# These MUST be added in your main PowerShell script that loads the XAML

# Account Type ComboBox Selection Changed
$AccountTypeComboBox.Add_SelectionChanged({
    try {
        if ($AccountTypeComboBox.SelectedItem -ne $null) {
            $global:SelectedAccountType = $AccountTypeComboBox.SelectedItem.Content.ToString().Trim()
        }
    }
    catch {
        # Silent fail
    }
})

# PC Farm ComboBox is now disabled/removed since all users go to the same group
# You should remove or disable this ComboBox from your XAML form
# $PCFarmComboBox.Add_SelectionChanged({
#     # This handler is no longer needed
# })

# ===== REMOTE ACCESS SUBMIT BUTTON =====
$RemoteAccessSubmitButton.Add_Click({

    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    # --- Get Values ---
    $computerName = $ComputerNameTextBox.Text.Trim()
    $userLANID    = $UserLANIDTextBox.Text.Trim()
    
    # Use the account type for informational purposes only
    $accountType = $global:SelectedAccountType

    # Debug output
    $outputTextBox.AppendText("=================================`n")
    $outputTextBox.AppendText("REMOTE ACCESS REQUEST`n")
    $outputTextBox.AppendText("=================================`n")
    $outputTextBox.AppendText("VALUES USED:`n")
    $outputTextBox.AppendText("  Computer Name: '$computerName'`n")
    $outputTextBox.AppendText("  User LAN ID: '$userLANID'`n")
    $outputTextBox.AppendText("  Account Type: '$accountType'`n")
    $outputTextBox.AppendText("  RDP Group: 'gs-DSS_OTI_RA_RDPUsers' (Hardcoded for ALL users)`n")
    $outputTextBox.AppendText("=================================`n`n")

    # --- Validation ---
    if (-not $computerName) {
        $outputTextBox.AppendText("ERROR: Computer Name is required.`n")
        return
    }
    
    if (-not $userLANID) {
        $outputTextBox.AppendText("ERROR: User LAN ID is required.`n")
        return
    }

    try {
        # --- HARDCODED RDP GROUP FOR ALL USERS ---
        $rdpGroup = "gs-DSS_OTI_RA_RDPUsers"
        $outputTextBox.AppendText("INFO: ALL users now go to hardcoded RDP group: $rdpGroup`n")

        # --- GET USER OBJECT ---
        $outputTextBox.AppendText("Looking up user: $userLANID...`n")
        $user = Get-ADUser -Identity $userLANID -Properties memberOf, extensionAttribute2, otherLoginWorkstations -ErrorAction Stop
        $outputTextBox.AppendText("User found: $($user.Name)`n")
        
        $groupDN = (Get-ADGroup -Identity $rdpGroup -ErrorAction Stop).DistinguishedName
        $outputTextBox.AppendText("Group found: $rdpGroup`n")

        # --- CHECK AND CLEAR extensionAttribute2 (NEW REQUIREMENT) ---
        if (-not [string]::IsNullOrEmpty($user.extensionAttribute2)) {
            $outputTextBox.AppendText("Found extensionAttribute2: '$($user.extensionAttribute2)'`n")
            Set-ADUser -Identity $userLANID -Clear extensionAttribute2 -ErrorAction Stop
            $outputTextBox.AppendText("CLEARED: extensionAttribute2 has been removed (no longer used).`n")
        } else {
            $outputTextBox.AppendText("extensionAttribute2: No value found (already clear).`n")
        }

        # --- ADD USER TO RDP GROUP ---
        if ($user.memberOf -notcontains $groupDN) {
            Add-ADGroupMember -Identity $rdpGroup -Members $userLANID -ErrorAction Stop
            $outputTextBox.AppendText("SUCCESS: User added to $rdpGroup.`n")
        }
        else {
            $outputTextBox.AppendText("INFO: User already in $rdpGroup.`n")
        }

        # --- REMOVE UNWANTED GROUPS ---
        $outputTextBox.AppendText("`nRemoving unwanted MFA groups...`n")
        $groupsToRemove = @(
            "GS-MFA-IPADusers",
            "GS-MFA-MACusers",
            "GS-MFA-NewRadiusAuthentication"
        )

        $removedGroups = @()
        foreach ($group in $groupsToRemove) {
            try {
                $groupObj = Get-ADGroup -Identity $group -ErrorAction Stop
                if ($user.memberOf -contains $groupObj.DistinguishedName) {
                    Remove-ADGroupMember -Identity $group -Members $userLANID -Confirm:$false -ErrorAction Stop
                    $removedGroups += $group
                    $outputTextBox.AppendText("Removed from $group.`n")
                }
                else {
                    $outputTextBox.AppendText("Not a member of $group.`n")
                }
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $outputTextBox.AppendText("Group $group does not exist (skipped).`n")
            }
            catch {
                $outputTextBox.AppendText("Warning: Failed to process $group - $($_.Exception.Message)`n")
            }
        }

        if ($removedGroups.Count -gt 0) {
            $outputTextBox.AppendText("SUMMARY: User removed from: $($removedGroups -join ', ').`n")
        } else {
            $outputTextBox.AppendText("SUMMARY: No groups removed.`n")
        }

        # --- COMPUTER CHECK ---
        $outputTextBox.AppendText("`nChecking computer: $computerName...`n")
        $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop
        $outputTextBox.AppendText("Computer found in AD.`n")

        if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
            $outputTextBox.AppendText("The computer $computerName is ONLINE.`n")

            # --- LOCAL RDP GROUP ON COMPUTER ---
            $outputTextBox.AppendText("Adding user to Remote Desktop Users group on computer...`n")
            try {
                Invoke-Command -ComputerName $computerName -ScriptBlock {
                    $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                    $members = @($group.PSBase.Invoke("Members")) | ForEach-Object {
                        $_.GetType().InvokeMember("ADsPath",'GetProperty',$null,$_,$null)
                    }

                    # Remove all existing members
                    foreach ($member in $members) { 
                        try {
                            $group.Remove($member)
                        } catch {
                            # Ignore errors when removing
                        }
                    }

                    # Add the current user
                    $group.Add("WinNT://$using:userLANID,user")
                }
                $outputTextBox.AppendText("User added to Remote Desktop Users group on $computerName.`n")
            }
            catch {
                $outputTextBox.AppendText("Warning: Could not modify local RDP group: $($_.Exception.Message)`n")
                $outputTextBox.AppendText("This may require manual intervention or admin rights on the target computer.`n")
            }

            # --- SET otherLoginWorkstations FOR ALL USERS (NEW REQUIREMENT) ---
            $outputTextBox.AppendText("Setting otherLoginWorkstations to $computerName...`n")
            Set-ADUser -Identity $userLANID -Replace @{ otherLoginWorkstations = $computerName }
            $outputTextBox.AppendText("SUCCESS: otherLoginWorkstations set to $computerName for ALL users.`n")

            # --- SUCCESS MESSAGE ---
            $outputTextBox.AppendText("`n" + ("="*50) + "`n")
            $outputTextBox.AppendText("REQUEST SUCCESSFULLY COMPLETED`n")
            $outputTextBox.AppendText("="*50 + "`n`n")
            $outputTextBox.AppendText(@"
PROCESS SUMMARY:
1. Cleared extensionAttribute2 (no longer used)
2. User added to gs-DSS_OTI_RA_RDPUsers (ALL users now)
3. Removed from unwanted MFA groups
4. Added to local Remote Desktop Users on $computerName
5. otherLoginWorkstations set to $computerName (for ALL users)

Please allow 30 minutes for permissions to fully propagate.

ACCESS INSTRUCTIONS:
1. Sign into https://myapplications.microsoft.com/
2. Click on the "My HRA PC" Tile
3. Sign in again with your credentials

Computer: $computerName
User: $userLANID ($($user.Name))
Account Type: $accountType
RDP Group: gs-DSS_OTI_RA_RDPUsers (All users)
"@)
        }
        else {
            $outputTextBox.AppendText("The computer $computerName is OFFLINE.`n")
            
            # --- SET otherLoginWorkstations FOR ALL USERS EVEN IF OFFLINE (NEW REQUIREMENT) ---
            $outputTextBox.AppendText("Setting otherLoginWorkstations to $computerName (computer offline)...`n")
            Set-ADUser -Identity $userLANID -Replace @{ otherLoginWorkstations = $computerName }
            $outputTextBox.AppendText("otherLoginWorkstations set to $computerName.`n")
            
            $outputTextBox.AppendText("`nACTION REQUIRED:")
            $outputTextBox.AppendText("`nThe computer must be online to complete local RDP configuration.")
            $outputTextBox.AppendText("`nAD group membership and otherLoginWorkstations have been updated, but local computer changes are pending.")
        }
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $outputTextBox.AppendText("ERROR: User or Computer not found in Active Directory.`n")
        $outputTextBox.AppendText("Please verify: $($_.Exception.Message)`n")
    }
    catch [System.UnauthorizedAccessException] {
        $outputTextBox.AppendText("ERROR: Permission denied.`n")
        $outputTextBox.AppendText("You may not have sufficient AD permissions for this operation.`n")
    }
    catch {
        $outputTextBox.AppendText("ERROR: $($_.Exception.Message)`n")
        $outputTextBox.AppendText("Exception Type: $($_.Exception.GetType().FullName)`n")
        
        # Clear extension attribute on error
        try {
            Set-ADUser -Identity $userLANID -Clear extensionAttribute2 -ErrorAction SilentlyContinue
            $outputTextBox.AppendText("Extension attribute cleared due to error.`n")
        }
        catch {
            $outputTextBox.AppendText("Could not clear extension attribute.`n")
        }
    }
})

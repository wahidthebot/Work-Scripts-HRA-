# Event handler for Remote Access Submit button
$remoteAccessSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $computerName   = $window.FindName("ComputerNameTextBox").Text
    $userLANID      = $window.FindName("UserLANIDTextBox").Text
    $groupSelection = $window.FindName("GroupSelectionComboBox").SelectedItem.Content

    if ([string]::IsNullOrEmpty($computerName) -or
        [string]::IsNullOrEmpty($userLANID) -or
        [string]::IsNullOrEmpty($groupSelection)) {

        $outputTextBox.AppendText(
            "Error: Computer Name, User LAN ID, and Group Selection are required.`n"
        )
        return
    }

    # Map the selection to the actual group name
    $groupToAdd = switch ($groupSelection) {
        "1" { "LUW-HRAPersonalRDPUsers" }
        "2" { "HRARDPUsers2" }
        "3" { "HRARDPUsers3" }
        default { "HRARDPUsers3" }
    }

    try {
        # --- ADD USER TO SELECTED GROUP ---
        $user = Get-ADUser -Identity $userLANID -Properties memberOf -ErrorAction Stop
        $groupDN = (Get-ADGroup -Identity $groupToAdd -ErrorAction Stop).DistinguishedName

        if ($user.memberOf -notcontains $groupDN) {
            Add-ADGroupMember -Identity $groupToAdd -Members $userLANID -ErrorAction Stop
            $outputTextBox.AppendText("SUCCESS: User added to $groupToAdd.`n")
        }
        else {
            $outputTextBox.AppendText("INFO: User is already in $groupToAdd.`n")
        }

        # --- REMOVE USER FROM UNWANTED GROUPS ---
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
                    Remove-ADGroupMember `
                        -Identity $group `
                        -Members $userLANID `
                        -Confirm:$false `
                        -ErrorAction Stop

                    $removedGroups += $group
                    $outputTextBox.AppendText("SUCCESS: User removed from $group.`n")
                }
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $outputTextBox.AppendText("INFO: Group $group does not exist (skipped).`n")
            }
            catch {
                $outputTextBox.AppendText(
                    "ERROR: Failed to process $group - $($_.Exception.Message)`n"
                )
            }
        }

        if ($removedGroups.Count -gt 0) {
            $outputTextBox.AppendText(
                "SUMMARY: User was removed from: $($removedGroups -join ', ').`n"
            )
        }
        else {
            $outputTextBox.AppendText(
                "SUMMARY: No groups were removed (user was not a member or groups didn't exist).`n"
            )
        }

        # --- COMPUTER CHECK ---
        $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop

        if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
            $outputTextBox.AppendText("The computer $computerName is online.`n")

            Invoke-Command -ComputerName $computerName -ScriptBlock {
                $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                $members = @(
                    $group.PSBase.Invoke("Members")
                ) | ForEach-Object {
                    $_.GetType().InvokeMember(
                        "ADsPath",
                        'GetProperty',
                        $null,
                        $_,
                        $null
                    )
                }

                foreach ($member in $members) {
                    $group.Remove($member)
                }

                $group.Add("WinNT://$using:userLANID,user")
            }

            $outputTextBox.AppendText(
                "The user $userLANID has been added to the Remote Desktop Users group on $computerName.`n"
            )

            Set-ADUser `
                -Identity $userLANID `
                -Replace @{ extensionAttribute2 = "$computerName.windows.nyc.hra.nycnet" }

            $outputTextBox.AppendText(
                "Close Out Ticket: User has been given access to PC.`n"
            )
        }
        else {
            $outputTextBox.AppendText(
                "The computer $computerName is found but offline.`n"
            )

            Set-ADUser -Identity $userLANID -Clear extensionAttribute2
            $outputTextBox.AppendText("Extension attribute has been removed.`n")
            $outputTextBox.AppendText(
                "Close Out Ticket: PC seems to be offline. Please check the connection of the PC.`n"
            )
        }
    }
    catch {
        $outputTextBox.AppendText(
            "The computer $computerName does not exist.`n"
        )

        Set-ADUser -Identity $userLANID -Clear extensionAttribute2
        $outputTextBox.AppendText("Extension attribute has been removed.`n")
        $outputTextBox.AppendText(
            "Close Out Ticket: PC was not able to be located. Please verify the PC name.`n"
        )
    }
})

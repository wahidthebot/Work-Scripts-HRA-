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


# Event handler for Group Management Submit button
$groupManagementSubmitButton.Add_Click({
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Script running...`n")
    $window.Dispatcher.Invoke([action]{}, "Render")  # Force UI update

    $userLanIDs = $window.FindName("GroupManagementUserLANIDTextBox").Text
    $groupNames = $window.FindName("GroupManagementGroupNamesTextBox").Text
    $action = $window.FindName("GroupManagementActionComboBox").SelectedItem.Content

    if ([string]::IsNullOrEmpty($userLanIDs) -or [string]::IsNullOrEmpty($groupNames)) {
        $outputTextBox.AppendText("❌ Error: Both User LAN ID(s) and Group Name(s) are required.`n")
        return
    }

    $userLanIDsArray = $userLanIDs -split "," | ForEach-Object { $_.Trim() }
    $groupNamesArray = $groupNames -split "," | ForEach-Object { $_.Trim() }

    foreach ($userLanID in $userLanIDsArray) {
        if ($action -eq "Add") {
            Add-UserToGroups -UserLanID $userLanID -GroupNames $groupNamesArray
        } elseif ($action -eq "Remove") {
            Remove-UserFromGroups -UserLanID $userLanID -GroupNames $groupNamesArray
        } else {
            $outputTextBox.AppendText("❌ Invalid action selected.`n")
        }
    }
})


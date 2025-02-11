<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Import Active Directory module
Import-Module ActiveDirectory

# Function to add a user to one or multiple groups
function Add-UserToGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserLanID,
        [Parameter(Mandatory = $true)]
        [string[]]$GroupNames
    )

    # Get the user's Distinguished Name (DN) based on their LAN ID
    $User = Get-ADUser -Filter "SamAccountName -eq '$UserLanID'" -Properties SamAccountName, MemberOf
    if ($null -eq $User) {
        Write-Output "User with LAN ID '$UserLanID' not found in Active Directory."
        return
    } else {
        Write-Output "User '$UserLanID' found."
    }

    foreach ($GroupName in $GroupNames) {
        # Check if the group exists
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -Properties Name
        if ($null -eq $Group) {
            Write-Output "Group '$GroupName' does not exist in Active Directory."
            continue
        } else {
            Write-Output "Group '$GroupName' found."
        }

        # Check if the user is already a member of the group
        $IsMember = $User.MemberOf -contains ("CN=$GroupName,OU=Groups,DC=YourDomain,DC=com")
        if ($IsMember) {
            Write-Output "User '$UserLanID' is already a member of the group '$GroupName'."
        } else {
            # Add the user to the group
            Add-ADGroupMember -Identity $GroupName -Members $User.DistinguishedName
            Write-Output "User '$UserLanID' has been added to the group '$GroupName'."
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

    # Get the user's Distinguished Name (DN) based on their LAN ID
    $User = Get-ADUser -Filter "SamAccountName -eq '$UserLanID'" -Properties SamAccountName
    if ($null -eq $User) {
        Write-Output "User with LAN ID '$UserLanID' not found in Active Directory."
        return
    } else {
        Write-Output "User '$UserLanID' found."
    }

    foreach ($GroupName in $GroupNames) {
        # Check if the group exists
        $Group = Get-ADGroup -Filter "Name -eq '$GroupName'" -Properties Name
        if ($null -eq $Group) {
            Write-Output "Group '$GroupName' does not exist in Active Directory."
            continue
        } else {
            Write-Output "Group '$GroupName' found."
        }

        # Check if the user is a member of the group
        $IsMember = Get-ADGroupMember -Identity $GroupName | Where-Object { $_.SamAccountName -eq $UserLanID }
        if ($IsMember) {
            # Remove the user from the group
            Remove-ADGroupMember -Identity $GroupName -Members $User.DistinguishedName -Confirm:$false
            Write-Output "User '$UserLanID' has been removed from the group '$GroupName'."
        } else {
            Write-Output "User '$UserLanID' is no longer a member of the group '$GroupName'."
        }
    }
}

# Main logic to prompt the user for action
$Action = Read-Host "Do you want to 'add' or 'remove' users from groups? (Enter 'add' or 'remove')"

if ($Action -eq "add") {
    $LanIDs = Read-Host "Enter the LAN ID(s) of the user(s) (comma-separated for multiple users)"
    $GroupNames = Read-Host "Enter the Active Directory group name(s) (comma-separated for multiple groups)"
    $LanIDsArray = $LanIDs -split ","
    $GroupNamesArray = $GroupNames -split ","

    foreach ($LanID in $LanIDsArray) {
        Add-UserToGroups -UserLanID $LanID.Trim() -GroupNames $GroupNamesArray
    }
} elseif ($Action -eq "remove") {
    $LanIDs = Read-Host "Enter the LAN ID(s) of the user(s) (comma-separated for multiple users)"
    $GroupNames = Read-Host "Enter the Active Directory group name(s) (comma-separated for multiple groups)"
    $LanIDsArray = $LanIDs -split ","
    $GroupNamesArray = $GroupNames -split ","

    foreach ($LanID in $LanIDsArray) {
        Remove-UserFromGroups -UserLanID $LanID.Trim() -GroupNames $GroupNamesArray
    }
} else {
    Write-Output "Invalid action. Please run the script again and enter 'add' or 'remove'."
}

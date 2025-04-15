<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Function to get user input
function Get-UserInput {
    param (
        [string]$prompt
    )
    Write-Host $prompt -NoNewline
    return Read-Host
}

# Get computer name and user LAN ID from the user
$computerName = Get-UserInput "Please enter the computer name: "
$userLANID = Get-UserInput "Please enter the user LAN ID: "

# Get group selection from user
Write-Host "`nSelect which HRARDPUsers group to add the user to:"
Write-Host "1 - HRARDPUsers1"
Write-Host "2 - HRARDPUsers2"
Write-Host "3 - HRARDPUsers3"
$groupSelection = Get-UserInput "Enter your choice (1-3): "

# Map selection to group name
$groupToAdd = switch ($groupSelection) {
    "1" { "HRARDPUsers1" }
    "2" { "HRARDPUsers2" }
    "3" { "HRARDPUsers3" }
    default { 
        Write-Host "Invalid selection. Defaulting to HRARDPUsers3."
        "HRARDPUsers3" 
    }
}

# --- GROUP MANAGEMENT SECTION ---
try {
    # 1. ADD USER TO SELECTED GROUP (IF NOT ALREADY A MEMBER)
    $user = Get-ADUser -Identity $userLANID -Properties memberof -ErrorAction Stop
    
    # Check if user is already in the selected group
    $groupDN = (Get-ADGroup -Identity $groupToAdd -ErrorAction Stop).DistinguishedName
    if (-not ($user.memberof -contains $groupDN)) {
        Add-ADGroupMember -Identity $groupToAdd -Members $userLANID -ErrorAction Stop
        Write-Host "SUCCESS: User added to $groupToAdd."
    } else {
        Write-Host "INFO: User is already in $groupToAdd."
    }

    # 2. REMOVE USER FROM UNWANTED MFA GROUPS
    $groupsToRemove = @("GS-MFA-IPADusers", "GS-MFA-MACusers", "GS-MFA-NewRadiusAuthentication")
    $removedGroups = @()

    foreach ($group in $groupsToRemove) {
        try {
            $groupObj = Get-ADGroup -Identity $group -ErrorAction Stop
            if ($user.memberof -contains $groupObj.DistinguishedName) {
                Remove-ADGroupMember -Identity $group -Members $userLANID -Confirm:$false -ErrorAction Stop
                $removedGroups += $group
                Write-Host "SUCCESS: User removed from $group."
            }
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Host "INFO: Group $group does not exist (skipped)."
        } catch {
            Write-Host "WARNING: Failed to process $group - $($_.Exception.Message)"
        }
    }

    # Summary of group changes
    if ($removedGroups.Count -gt 0) {
        Write-Host "SUMMARY: User was removed from: $($removedGroups -join ', ')."
    }
}
catch {
    Write-Host "ERROR: Group management failed - $($_.Exception.Message)"
    exit
}

# --- COMPUTER MANAGEMENT SECTION ---
try {
    $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop
    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
        Write-Host "The computer $computerName is online."
        
        # Remove current users from Remote Desktop Users group and add the specified user
        try {
            Invoke-Command -ComputerName $computerName -ScriptBlock {
                $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                $members = @($group.PSBase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $_, $null) }
                foreach ($member in $members) {
                    $group.Remove($member)
                }
                $group.Add("WinNT://$using:userLANID,user")
            }
            Write-Host "SUCCESS: $userLANID added to Remote Desktop Users on $computerName."
            
            # Update extensionAttribute2
            Set-ADUser -Identity $userLANID -Replace @{extensionAttribute2 = "$computerName.windows.nyc.hra.nycnet"}
            Write-Host "SUCCESS: extensionAttribute2 updated."
            
            Write-Host "Close Out Ticket: User has been given access to PC."
        } catch {
            Write-Host "ERROR: Remote Desktop configuration failed - $($_.Exception.Message)"
        }
    } else {
        Write-Host "WARNING: Computer $computerName is found but offline."
        Set-ADUser -Identity $userLANID -Clear extensionAttribute2
        Write-Host "INFO: extensionAttribute2 cleared (computer offline)."
        Write-Host "Close Out Ticket: PC seems to be offline. Please check the connection."
    }
} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    Write-Host "ERROR: Computer $computerName does not exist in AD."
    Set-ADUser -Identity $userLANID -Clear extensionAttribute2
    Write-Host "INFO: extensionAttribute2 cleared (computer not found)."
    Write-Host "Close Out Ticket: PC was not able to be located. Please verify the PC name."
} catch {
    Write-Host "ERROR: Computer verification failed - $($_.Exception.Message)"
}

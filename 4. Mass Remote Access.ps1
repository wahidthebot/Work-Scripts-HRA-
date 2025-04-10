<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Path to the CSV file
$csvFilePath = "C:\path\to\your\file.csv"  # <-- Update this path

# Process each entry in the CSV file
Import-Csv -Path $csvFilePath | ForEach-Object {
    $userLANID = $_.lanID
    $computerName = $_.PC

    if ($userLANID -and $computerName) {
        Write-Host "`nProcessing user $userLANID for computer $computerName"
        
        # --- GROUP MANAGEMENT SECTION ---
        try {
            $user = Get-ADUser -Identity $userLANID -Properties memberof -ErrorAction Stop

            # 1. ADD USER TO HRARDPUsers3 (IF NOT ALREADY A MEMBER)
            $groupToAdd = "HRARDPUsers3"
            $groupDN = (Get-ADGroup -Identity $groupToAdd -ErrorAction Stop).DistinguishedName
            if (-not ($user.memberof -contains $groupDN)) {
                Add-ADGroupMember -Identity $groupToAdd -Members $userLANID -ErrorAction Stop
                Write-Host "SUCCESS: User added to $groupToAdd"
            } else {
                Write-Host "INFO: User is already in $groupToAdd"
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
                        Write-Host "SUCCESS: User removed from $group"
                    }
                } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                    Write-Host "INFO: Group $group does not exist (skipped)"
                } catch {
                    Write-Host "WARNING: Failed to process $group - $($_.Exception.Message)"
                }
            }

            if ($removedGroups.Count -gt 0) {
                Write-Host "SUMMARY: User was removed from: $($removedGroups -join ', ')"
            }
        } catch {
            Write-Host "ERROR: Group management failed - $($_.Exception.Message)"
            continue
        }

        # --- COMPUTER MANAGEMENT SECTION ---
        try {
            $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop
            if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
                Write-Host "STATUS: Computer $computerName is online"
                
                try {
                    # Configure Remote Desktop access
                    Invoke-Command -ComputerName $computerName -ScriptBlock {
                        $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                        $members = @($group.PSBase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $_, $null) }
                        foreach ($member in $members) {
                            $group.Remove($member)
                        }
                        $group.Add("WinNT://$using:userLANID,user")
                    }
                    Write-Host "SUCCESS: $userLANID added to Remote Desktop Users"
                    
                    # Update extension attribute
                    Set-ADUser -Identity $userLANID -Replace @{extensionAttribute2 = "$computerName.windows.nyc.hra.nycnet"}
                    Write-Host "SUCCESS: extensionAttribute2 updated"
                    
                    Write-Host "CLOSE OUT: User $userLANID has been given access to $computerName"
                } catch {
                    Write-Host "ERROR: Remote Desktop configuration failed - $($_.Exception.Message)"
                }
            } else {
                Write-Host "WARNING: Computer $computerName is offline"
                Set-ADUser -Identity $userLANID -Clear extensionAttribute2
                Write-Host "INFO: extensionAttribute2 cleared (computer offline)"
                Write-Host "CLOSE OUT: $computerName is offline - please check connection"
            }
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Host "ERROR: Computer $computerName not found in AD"
            Set-ADUser -Identity $userLANID -Clear extensionAttribute2
            Write-Host "INFO: extensionAttribute2 cleared (computer not found)"
            Write-Host "CLOSE OUT: $computerName could not be located - verify PC name"
        } catch {
            Write-Host "ERROR: Computer verification failed - $($_.Exception.Message)"
        }
    } else {
        Write-Host "ERROR: Invalid CSV entry - both lanID and PC fields are required"
    }
}

<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Path to the CSV file
$csvFilePath = ""

# Verify if user is a member of specific groups and add if not
$groups = @("GS-MFA-IPADusers", "GS-MFA-MACusers", "GS-MFA-NewRadiusAuthentication")

# Process each entry in the CSV file
Import-Csv -Path $csvFilePath | ForEach-Object {
    $userLANID = $_.lanID
    $computerName = $_.PC
    $notMembers = @()

    if ($userLANID -and $computerName) {
        foreach ($group in $groups) {
            $groupDN = (Get-ADGroup -Identity $group).DistinguishedName
            if (-not (Get-ADUser -Identity $userLANID -Properties memberof | Where-Object { $_.memberof -contains $groupDN })) {
                Add-ADGroupMember -Identity $group -Members $userLANID
                $notMembers += $group
            }
        }

        # Inform the user about the groups they were added to
        if ($notMembers.Count -gt 0) {
            Write-Host "The user $userLANID was not a part of the following groups and has been added: " -NoNewline
            $notMembers -join ", "
        } else {
            Write-Host "The user $userLANID is already a member of all the required groups."
        }

        # Check if the computer exists and ping to check if it's online
        try {
            $computer = Get-ADComputer -Identity $computerName -ErrorAction Stop
            if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
                Write-Host "The computer $computerName is online."
                
                # Remove current users from Remote Desktop Users group and add the specified user
                Invoke-Command -ComputerName $computerName -ScriptBlock {
                    $group = [ADSI]"WinNT://./Remote Desktop Users,group"
                    $members = @($group.PSBase.Invoke("Members")) | ForEach-Object { $_.GetType().InvokeMember("ADsPath", 'GetProperty', $null, $_, $null) }
                    foreach ($member in $members) {
                        $group.Remove($member)
                    }
                    $group.Add("WinNT://$using:userLANID,user")
                }
                Write-Host "The user $userLANID has been added to the Remote Desktop Users group on $computerName."
                
                # Update extensionAttribute2 for the user
                $user = Get-ADUser -Identity $userLANID -Properties extensionAttribute2
                $user.extensionAttribute2 = "$computerName.windows.nyc.hra.nycnet"
                Set-ADUser -Identity $userLANID -Replace @{extensionAttribute2 = $user.extensionAttribute2}
                
                # Close out ticket
                Write-Host "Close Out Ticket: User $userLANID has been given access to PC $computerName."
            } else {
                Write-Host "The computer $computerName is found but offline."
                
                # Remove extensionAttribute2 for the user if the computer is offline
                Set-ADUser -Identity $userLANID -Clear extensionAttribute2
                Write-Host "Extension attribute has been removed for $userLANID."
                
                # Close out ticket
                Write-Host "Close Out Ticket: PC $computerName seems to be offline. Please check the connection of the PC."
            }
        } catch {
            Write-Host "The computer $computerName does not exist."
            
            # Remove extensionAttribute2 for the user if the computer does not exist
            Set-ADUser -Identity $userLANID -Clear extensionAttribute2
            Write-Host "Extension attribute has been removed for $userLANID."
            
            # Close out ticket
            Write-Host "Close Out Ticket: PC $computerName was not able to be located. Please verify the PC name."
        }
    } else {
        Write-Host "Invalid entry in CSV file. Ensure 'lanID' and 'PC' are not empty."
    }
}

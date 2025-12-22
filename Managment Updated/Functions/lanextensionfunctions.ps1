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

        # Find the user in Active Directory
        $user = Get-ADUser -Filter {SamAccountName -eq $userLanId} -Properties AccountExpirationDate, UserPrincipalName, Enabled, Description

        if ($null -eq $user) {
            $outputTextBox.AppendText("User not found.`n")
            return
        }

        # Enable account if disabled
        if ($user.Enabled -eq $false) {
            Set-ADUser -Identity $user -Enabled $true
            $outputTextBox.AppendText("User account was disabled and has now been enabled.`n")
        } else {
            $outputTextBox.AppendText("User account is enabled.`n")
        }

        # Ensure "Account never expires" is unticked
        if ($user.AccountExpirationDate -eq $null) {
            Set-ADUser -Identity $user -AccountExpirationDate (Get-Date -Date "12/31/9999") # Temporary date to untick the box
        }

        # Calculate the new extension date
        $newExtendDate = (Get-Date -Date $extendDate).AddDays(1)

        # --- UPDATE EXPIRATION AND DESCRIPTION FIRST ---
        Set-ADUser -Identity $user -AccountExpirationDate $newExtendDate

        $newDescription = "Extended as per $ticketNumber $initials | "
        $existingDescription = $user.Description
        Set-ADUser -Identity $user -Description "$newDescription$existingDescription"

        $outputTextBox.AppendText("User's account has been extended to $extendDate.`n")

        # --- MOVE THE OBJECT LAST ---
        if (-not $isDHS) {
            $newOU = "OU=Temps (Replacing 15 MTC Temps OU),OU=470 Vanderbilt,OU=People,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $newOU
            $outputTextBox.AppendText("User has been moved to 470 Vanderbilt Temps OU.`n")
        } else {
            $dhsOU = "OU=People,OU=DHS Resources,OU=DHS,DC=windows,DC=nyc,DC=hra,DC=nycnet"
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $dhsOU
            $outputTextBox.AppendText("User is DHS. Account moved to DHS OU.`n")
        }

    } catch {
        $outputTextBox.AppendText("Error: $_`n")
    } finally {
        # Clear the fields after submission (except Initials)
        $window.FindName("LanExtensionUserLANIDTextBox").Text = ""
        $window.FindName("LanExtensionDateTextBox").Text = ""
        $window.FindName("LanExtensionTicketNumberTextBox").Text = ""
    }
})
function Format-PhoneNumber {
    param (
        [string]$PhoneNumber
    )
    # Remove non-numeric characters except plus
    $formattedNumber = $PhoneNumber -replace '[^\d\+]', ''
    if (-not $formattedNumber.StartsWith("+1")) {
        $formattedNumber = "+1" + $formattedNumber.TrimStart('+', '1')
    }
    return $formattedNumber
}

function Update-AuthenticationPhoneNumber {
    param (
        [string]$UserEmail,
        [string]$PhoneNumber,
        [string]$MethodType
    )

    try {
        $user = Get-MgUser -Filter "mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'"
        if (-not $user) {
            $outputTextBox.AppendText("Error: User not found.`n")
            return
        }

        $formattedPhoneNumber = Format-PhoneNumber -PhoneNumber $PhoneNumber

        # Normalize method type
        switch ($MethodType) {
            "Mobile" { $MethodType = "mobile" }
            "Alternate Mobile" { $MethodType = "alternateMobile" }
            "Office" { $MethodType = "office" }
            default { $MethodType = $MethodType.ToLower() }
        }

        $existingMethods = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id

        # Update or add phone number
        try {
            $existingPhoneMethod = $existingMethods | Where-Object { $_.PhoneType -eq $MethodType }
            if ($existingPhoneMethod) {
                Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber -ErrorAction Stop
                $outputTextBox.AppendText("User phone number has been changed.`n")
            } else {
                New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType $MethodType -ErrorAction Stop
                $outputTextBox.AppendText("User phone number has been added.`n")
            }
        } catch {
            $errorMessage = $_.Exception.Message
            $outputTextBox.AppendText("Error: $errorMessage`n")
            $outputTextBox.AppendText("Please activate your roles in Azure AD and try again.`n")
        }

    } catch {
        $outputTextBox.AppendText("Error: Could not retrieve user: $_`n")
    }
}

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
        Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All"
        Update-AuthenticationPhoneNumber -UserEmail $userEmail -PhoneNumber $phoneNumber -MethodType $methodType
    } catch {
        $outputTextBox.AppendText("Error: Failed to connect to Microsoft Graph or execute operation.`n")
        $outputTextBox.AppendText("Details: $_`n")
    }
})

<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Connect Microsoft Graph
Connect-MgGraph -Scopes "UserAuthenticationMethod.ReadWrite.All"

# format phone number
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

# update phone numbers in the authentication methods section
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
            Write-Host "User not found!" -ForegroundColor Red
            return
        }
        
        # Format number
        $formattedPhoneNumber = Format-PhoneNumber -PhoneNumber $PhoneNumber

        # phone method type based on user input
        switch ($MethodType) {
            1 { $MethodType = "mobile" }
            2 { $MethodType = "alternateMobile" }
            3 { $MethodType = "office" }
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
                    Write-Host "Close out ticker: User phone number has been changed"
                } else {
                    # Add a new mobile phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "mobile"
                    Write-Host "Close out ticker: User phone number has been added"
                }
            }
            "alternatemobile" {
                # Check if the user already has an alternate phone number
                $existingPhoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id | Where-Object { $_.PhoneType -eq "alternateMobile" }
                if ($existingPhoneMethod) {
                    # Update the existing alternate phone number
                    Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber
                    Write-Host "Close out ticker: User phone number has been changed"
                } else {
                    # Add a new alternate phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "alternateMobile"
                    Write-Host "Close out ticker: User phone number has been added"
                }
            }
            "office" {
                # Check if the user already has an office phone number
                $existingPhoneMethod = Get-MgUserAuthenticationPhoneMethod -UserId $user.Id | Where-Object { $_.PhoneType -eq "office" }
                if ($existingPhoneMethod) {
                    # Update the existing office phone number
                    Update-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneAuthenticationMethodId $existingPhoneMethod.Id -PhoneNumber $formattedPhoneNumber
                    Write-Host "Close out ticker: User phone number has been changed"
                } else {
                    # Add a new office phone number
                    New-MgUserAuthenticationPhoneMethod -UserId $user.Id -PhoneNumber $formattedPhoneNumber -PhoneType "office"
                    Write-Host "Close out ticker: User phone number has been added"
                }
            }
            default {
                Write-Host "Invalid phone method type specified!" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Write-Host "Please activate your roles and try again." -ForegroundColor Yellow
    }
}

# Example usage
# Collect user input
$UserEmail = Read-Host "Enter the user's email"
$PhoneNumber = Read-Host "Enter the phone number"
$MethodTypeInput = Read-Host "Enter the phone method type (1 for mobile, 2 for alternate, 3 for office, or the text)"

# Call the function
Update-AuthenticationPhoneNumber -UserEmail $UserEmail -PhoneNumber $PhoneNumber -MethodType $MethodTypeInput

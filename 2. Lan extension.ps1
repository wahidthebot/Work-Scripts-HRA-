<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

# Prompt for user LAN ID, the date to extend to, and the ticket number
$userLanId = Read-Host "Enter User LAN ID"
$extendDate = Read-Host "Enter the date you want to extend to (MM/DD/YYYY)"
$ticketNumber = Read-Host "Enter the ticket number"

# Import the Active Directory module
Import-Module ActiveDirectory

# Find the user in Active Directory
$user = Get-ADUser -Filter {SamAccountName -eq $userLanId} -Properties AccountExpirationDate, UserPrincipalName, Enabled, Description

# Check if the user was found
if ($null -eq $user) {
    Write-Host "User not found"
    exit
}

# Check if the user's account is disabled and enable it if necessary
if ($user.Enabled -eq $false) {
    Set-ADUser -Identity $user -Enabled $true
    Write-Host "User account was disabled and has now been enabled."
} else {
    Write-Host "User account is enabled."
}

# Check if the "Account never expires" box is ticked
if ($user.AccountExpirationDate -eq $null) {
    # Untick the "Account never expires" box by setting an expiration date
    Set-ADUser -Identity $user -AccountExpirationDate (Get-Date -Date "12/31/9999") # Temporary date to untick the box
}

# Check if the extension date is more than one year ahead
$currentDate = (Get-Date).Date
$maxExtendDate = $currentDate.AddYears(1)

if ([datetime]::ParseExact($extendDate, 'MM/dd/yyyy', $null).Date -gt $maxExtendDate) {
    Write-Host "Cannot extend the account more than one year ahead in time"
    exit
}

# Calculate the new extension date by adding one day to the given date
$newExtendDate = (Get-Date -Date $extendDate).AddDays(1)

# Update the user's account extension date
Set-ADUser -Identity $user -AccountExpirationDate $newExtendDate

# Update the user's description field by prepending the new description
$newDescription = "Extended as per $ticketNumber | "
$existingDescription = $user.Description
Set-ADUser -Identity $user -Description "$newDescription$existingDescription"

Write-Host "User's account has been extended to $extendDate."
Write-Host "Ticket Close Test: User's account has been extended to $extendDate. Replace 'Wahid Hussain' with your own initials."


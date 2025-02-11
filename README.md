# Work-Scripts
Scripts Used for Resolving Tickets

MFA Update NUmber: The "MFA update number.ps1" script is designed to update a user's phone number in their authentication methods using the Microsoft Graph API. It performs tasks such as formatting phone numbers, retrieving user information, and updating or adding phone numbers based on the specified method type (mobile, alternate mobile, or office). The script ensures accurate phone number formatting and provides feedback on successful updates, while also handling errors and edge cases gracefully. This helps maintain up-to-date and correctly formatted phone numbers for Multi-Factor Authentication (MFA) purposes.

Remote Access: The "Remote access.ps1" script collects a computer name and user LAN ID, verifies the user's membership in specific AD groups, and adds them if they aren't members. It then checks if the specified computer is online, updates the user's Remote Desktop Users group membership, and modifies the user's extensionAttribute2 with the computer's information, providing appropriate feedback throughout the process.


Add User To Group (manual): This PowerShell script manages Active Directory (AD) group memberships for users. It allows administrators to add or remove users from specified AD groups based on their LAN ID. The script prompts for user input to specify the action (add or remove), the LAN IDs, and the group names. It checks for the existence of users and groups, adds users to the groups if they aren't already members, or removes them if they are members, providing feedback throughout the process.

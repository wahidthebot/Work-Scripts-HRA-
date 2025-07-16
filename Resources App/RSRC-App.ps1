<#
    Copyright (c) 2025 Wahid Hussain
    This script is licensed under the MIT License.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "HRA Active Directory Management Tool"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = "#f0f0f0"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(860, 630)
$tabControl.Anchor = "Top, Bottom, Left, Right"

# ===== TAB 1: GROUP OPERATIONS =====
$tabPage1 = New-Object System.Windows.Forms.TabPage
$tabPage1.Text = "Group Management"
$tabPage1.BackColor = "#ffffff"

# Export Group Members
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Location = New-Object System.Drawing.Point(20, 20)
$groupBox1.Size = New-Object System.Drawing.Size(800, 150)
$groupBox1.Text = "Export Group Members"
$groupBox1.ForeColor = "#0066cc"

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20, 30)
$label1.Size = New-Object System.Drawing.Size(150, 20)
$label1.Text = "Group Name:"
$groupBox1.Controls.Add($label1)

$txtGroupName = New-Object System.Windows.Forms.TextBox
$txtGroupName.Location = New-Object System.Drawing.Point(180, 30)
$txtGroupName.Size = New-Object System.Drawing.Size(200, 25)
$groupBox1.Controls.Add($txtGroupName)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20, 70)
$label2.Size = New-Object System.Drawing.Size(150, 20)
$label2.Text = "Output File Path:"
$groupBox1.Controls.Add($label2)

$txtOutputPath = New-Object System.Windows.Forms.TextBox
$txtOutputPath.Location = New-Object System.Drawing.Point(180, 70)
$txtOutputPath.Size = New-Object System.Drawing.Size(400, 25)
$txtOutputPath.Text = "C:\temp\group_members.csv"
$groupBox1.Controls.Add($txtOutputPath)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(180, 110)
$btnExport.Size = New-Object System.Drawing.Size(150, 30)
$btnExport.Text = "Export Members"
$btnExport.BackColor = "#0066cc"
$btnExport.ForeColor = "White"
$btnExport.Add_Click({
    try {
        Get-ADGroup -Identity $txtGroupName.Text | Get-ADGroupMember | 
            Select-Object name, samaccountname, mail, objectclass, distinguishedname | 
            Export-Csv $txtOutputPath.Text -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Group members exported successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox1.Controls.Add($btnExport)

# Add/Remove Group Members
$groupBox2 = New-Object System.Windows.Forms.GroupBox
$groupBox2.Location = New-Object System.Drawing.Point(20, 190)
$groupBox2.Size = New-Object System.Drawing.Size(800, 180)
$groupBox2.Text = "Modify Group Members"
$groupBox2.ForeColor = "#0066cc"

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(20, 30)
$label3.Size = New-Object System.Drawing.Size(150, 20)
$label3.Text = "Group Name:"
$groupBox2.Controls.Add($label3)

$txtTargetGroup = New-Object System.Windows.Forms.TextBox
$txtTargetGroup.Location = New-Object System.Drawing.Point(180, 30)
$txtTargetGroup.Size = New-Object System.Drawing.Size(200, 25)
$groupBox2.Controls.Add($txtTargetGroup)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(20, 70)
$label4.Size = New-Object System.Drawing.Size(150, 20)
$label4.Text = "User SAMs (comma-separated):"
$groupBox2.Controls.Add($label4)

$txtUserSAMs = New-Object System.Windows.Forms.TextBox
$txtUserSAMs.Location = New-Object System.Drawing.Point(180, 70)
$txtUserSAMs.Size = New-Object System.Drawing.Size(400, 50)
$txtUserSAMs.Multiline = $true
$groupBox2.Controls.Add($txtUserSAMs)

$btnAddMembers = New-Object System.Windows.Forms.Button
$btnAddMembers.Location = New-Object System.Drawing.Point(180, 130)
$btnAddMembers.Size = New-Object System.Drawing.Size(150, 30)
$btnAddMembers.Text = "Add Members"
$btnAddMembers.BackColor = "#4CAF50"
$btnAddMembers.ForeColor = "White"
$btnAddMembers.Add_Click({
    try {
        $members = $txtUserSAMs.Text -split ',' | ForEach-Object { $_.Trim() }
        Add-ADGroupMember -Identity $txtTargetGroup.Text -Members $members
        [System.Windows.Forms.MessageBox]::Show("Users added successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox2.Controls.Add($btnAddMembers)

$btnRemoveMembers = New-Object System.Windows.Forms.Button
$btnRemoveMembers.Location = New-Object System.Drawing.Point(350, 130)
$btnRemoveMembers.Size = New-Object System.Drawing.Size(150, 30)
$btnRemoveMembers.Text = "Remove Members"
$btnRemoveMembers.BackColor = "#f44336"
$btnRemoveMembers.ForeColor = "White"
$btnRemoveMembers.Add_Click({
    try {
        $members = $txtUserSAMs.Text -split ',' | ForEach-Object { $_.Trim() }
        Remove-ADGroupMember -Identity $txtTargetGroup.Text -Members $members -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Users removed successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox2.Controls.Add($btnRemoveMembers)

$tabPage1.Controls.Add($groupBox1)
$tabPage1.Controls.Add($groupBox2)

# ===== TAB 2: COMPUTER MANAGEMENT =====
$tabPage2 = New-Object System.Windows.Forms.TabPage
$tabPage2.Text = "Computer Management"
$tabPage2.BackColor = "#ffffff"

# Reboot Computers
$groupBox3 = New-Object System.Windows.Forms.GroupBox
$groupBox3.Location = New-Object System.Drawing.Point(20, 20)
$groupBox3.Size = New-Object System.Drawing.Size(800, 180)
$groupBox3.Text = "Reboot Computers"
$groupBox3.ForeColor = "#0066cc"

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(20, 30)
$label5.Size = New-Object System.Drawing.Size(200, 20)
$label5.Text = "Computer Names (one per line):"
$groupBox3.Controls.Add($label5)

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Location = New-Object System.Drawing.Point(20, 60)
$txtComputers.Size = New-Object System.Drawing.Size(400, 100)
$txtComputers.Multiline = $true
$txtComputers.ScrollBars = "Vertical"
$groupBox3.Controls.Add($txtComputers)

$btnReboot = New-Object System.Windows.Forms.Button
$btnReboot.Location = New-Object System.Drawing.Point(450, 60)
$btnReboot.Size = New-Object System.Drawing.Size(150, 30)
$btnReboot.Text = "Reboot Computers"
$btnReboot.BackColor = "#FF9800"
$btnReboot.ForeColor = "White"
$btnReboot.Add_Click({
    try {
        $computers = $txtComputers.Text -split "`r`n" | Where-Object { $_ -ne "" }
        foreach ($computer in $computers) {
            Restart-Computer -ComputerName $computer -Force
        }
        [System.Windows.Forms.MessageBox]::Show("Reboot commands sent!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox3.Controls.Add($btnReboot)

$tabPage2.Controls.Add($groupBox3)

# ===== TAB 3: USER MANAGEMENT =====
$tabPage3 = New-Object System.Windows.Forms.TabPage
$tabPage3.Text = "User Management"
$tabPage3.BackColor = "#ffffff"

# User Information
$groupBox4 = New-Object System.Windows.Forms.GroupBox
$groupBox4.Location = New-Object System.Drawing.Point(20, 20)
$groupBox4.Size = New-Object System.Drawing.Size(800, 150)
$groupBox4.Text = "User Information"
$groupBox4.ForeColor = "#0066cc"

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(20, 30)
$label6.Size = New-Object System.Drawing.Size(150, 20)
$label6.Text = "Username/SAM:"
$groupBox4.Controls.Add($label6)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(180, 30)
$txtUsername.Size = New-Object System.Drawing.Size(200, 25)
$groupBox4.Controls.Add($txtUsername)

$btnGetUserInfo = New-Object System.Windows.Forms.Button
$btnGetUserInfo.Location = New-Object System.Drawing.Point(400, 30)
$btnGetUserInfo.Size = New-Object System.Drawing.Size(150, 30)
$btnGetUserInfo.Text = "Get User Info"
$btnGetUserInfo.BackColor = "#9C27B0"
$btnGetUserInfo.ForeColor = "White"
$btnGetUserInfo.Add_Click({
    try {
        $output = net user $txtUsername.Text /domain
        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "User Information"
        $resultForm.Size = New-Object System.Drawing.Size(500, 400)
        
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Dock = "Fill"
        $textBox.Text = $output -join "`r`n"
        $textBox.ReadOnly = $true
        
        $resultForm.Controls.Add($textBox)
        $resultForm.ShowDialog()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox4.Controls.Add($btnGetUserInfo)

# Export Users from OU
$groupBox5 = New-Object System.Windows.Forms.GroupBox
$groupBox5.Location = New-Object System.Drawing.Point(20, 190)
$groupBox5.Size = New-Object System.Drawing.Size(800, 200)
$groupBox5.Text = "Export Users from OU"
$groupBox5.ForeColor = "#0066cc"

$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(20, 30)
$label7.Size = New-Object System.Drawing.Size(150, 20)
$label7.Text = "OU DistinguishedName:"
$groupBox5.Controls.Add($label7)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(180, 30)
$txtOU.Size = New-Object System.Drawing.Size(500, 25)
$txtOU.Text = "OU=31st Floor,OU=4WTC,OU=People,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
$groupBox5.Controls.Add($txtOU)

$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(20, 70)
$label8.Size = New-Object System.Drawing.Size(150, 20)
$label8.Text = "Output File Path:"
$groupBox5.Controls.Add($label8)

$txtOUOutputPath = New-Object System.Windows.Forms.TextBox
$txtOUOutputPath.Location = New-Object System.Drawing.Point(180, 70)
$txtOUOutputPath.Size = New-Object System.Drawing.Size(400, 25)
$txtOUOutputPath.Text = "C:\temp\ou_users.csv"
$groupBox5.Controls.Add($txtOUOutputPath)

$btnExportOU = New-Object System.Windows.Forms.Button
$btnExportOU.Location = New-Object System.Drawing.Point(180, 110)
$btnExportOU.Size = New-Object System.Drawing.Size(150, 30)
$btnExportOU.Text = "Export Users"
$btnExportOU.BackColor = "#0066cc"
$btnExportOU.ForeColor = "White"
$btnExportOU.Add_Click({
    try {
        Get-ADUser -Filter * -SearchBase $txtOU.Text | 
            Select-Object SamAccountName | 
            Export-Csv -Path $txtOUOutputPath.Text -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Users exported successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox5.Controls.Add($btnExportOU)

$tabPage3.Controls.Add($groupBox4)
$tabPage3.Controls.Add($groupBox5)

# ===== TAB 4: COMPUTER ADMINISTRATION =====
$tabPage4 = New-Object System.Windows.Forms.TabPage
$tabPage4.Text = "Computer Admin"
$tabPage4.BackColor = "#ffffff"

# Local Admin Management
$groupBox6 = New-Object System.Windows.Forms.GroupBox
$groupBox6.Location = New-Object System.Drawing.Point(20, 20)
$groupBox6.Size = New-Object System.Drawing.Size(800, 180)
$groupBox6.Text = "Local Administrator Management"
$groupBox6.ForeColor = "#0066cc"

$label9 = New-Object System.Windows.Forms.Label
$label9.Location = New-Object System.Drawing.Point(20, 30)
$label9.Size = New-Object System.Drawing.Size(150, 20)
$label9.Text = "Computer Name:"
$groupBox6.Controls.Add($label9)

$txtTargetComputer = New-Object System.Windows.Forms.TextBox
$txtTargetComputer.Location = New-Object System.Drawing.Point(180, 30)
$txtTargetComputer.Size = New-Object System.Drawing.Size(200, 25)
$txtTargetComputer.Text = "wavditspers64.windows.nyc.hra.nycnet"
$groupBox6.Controls.Add($txtTargetComputer)

$label10 = New-Object System.Windows.Forms.Label
$label10.Location = New-Object System.Drawing.Point(20, 70)
$label10.Size = New-Object System.Drawing.Size(150, 20)
$label10.Text = "Username to Add:"
$groupBox6.Controls.Add($label10)

$txtAdminUser = New-Object System.Windows.Forms.TextBox
$txtAdminUser.Location = New-Object System.Drawing.Point(180, 70)
$txtAdminUser.Size = New-Object System.Drawing.Size(200, 25)
$txtAdminUser.Text = "windows\ckesa0701"
$groupBox6.Controls.Add($txtAdminUser)

$btnAddAdmin = New-Object System.Windows.Forms.Button
$btnAddAdmin.Location = New-Object System.Drawing.Point(180, 110)
$btnAddAdmin.Size = New-Object System.Drawing.Size(200, 30)
$btnAddAdmin.Text = "Add Local Administrator"
$btnAddAdmin.BackColor = "#4CAF50"
$btnAddAdmin.ForeColor = "White"
$btnAddAdmin.Add_Click({
    try {
        $computer = $txtTargetComputer.Text
        $user = $txtAdminUser.Text
        
        Invoke-Command -ScriptBlock {
            param($user)
            net localgroup "administrators" /add $user
        } -ComputerName $computer -ArgumentList $user
        
        [System.Windows.Forms.MessageBox]::Show("User $user added as local admin on $computer", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox6.Controls.Add($btnAddAdmin)

$tabPage4.Controls.Add($groupBox6)

# Computer OU Management
$groupBox7 = New-Object System.Windows.Forms.GroupBox
$groupBox7.Location = New-Object System.Drawing.Point(20, 220)
$groupBox7.Size = New-Object System.Drawing.Size(800, 200)
$groupBox7.Text = "Computer OU Management"
$groupBox7.ForeColor = "#0066cc"

$label11 = New-Object System.Windows.Forms.Label
$label11.Location = New-Object System.Drawing.Point(20, 30)
$label11.Size = New-Object System.Drawing.Size(150, 20)
$label11.Text = "Computer Name:"
$groupBox7.Controls.Add($label11)

$txtMoveComputer = New-Object System.Windows.Forms.TextBox
$txtMoveComputer.Location = New-Object System.Drawing.Point(180, 30)
$txtMoveComputer.Size = New-Object System.Drawing.Size(200, 25)
$txtMoveComputer.Text = "w470van07j522"
$groupBox7.Controls.Add($txtMoveComputer)

$label12 = New-Object System.Windows.Forms.Label
$label12.Location = New-Object System.Drawing.Point(20, 70)
$label12.Size = New-Object System.Drawing.Size(150, 20)
$label12.Text = "Target OU:"
$groupBox7.Controls.Add($label12)

$txtTargetOU = New-Object System.Windows.Forms.TextBox
$txtTargetOU.Location = New-Object System.Drawing.Point(180, 70)
$txtTargetOU.Size = New-Object System.Drawing.Size(500, 25)
$txtTargetOU.Text = "OU=Locally Managed PCs,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
$groupBox7.Controls.Add($txtTargetOU)

$btnCheckComputer = New-Object System.Windows.Forms.Button
$btnCheckComputer.Location = New-Object System.Drawing.Point(180, 110)
$btnCheckComputer.Size = New-Object System.Drawing.Size(150, 30)
$btnCheckComputer.Text = "Check Computer"
$btnCheckComputer.BackColor = "#2196F3"
$btnCheckComputer.ForeColor = "White"
$btnCheckComputer.Add_Click({
    try {
        $computer = Get-ADComputer -Identity $txtMoveComputer.Text -Properties *
        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "Computer Information"
        $resultForm.Size = New-Object System.Drawing.Size(600, 400)
        
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Dock = "Fill"
        $textBox.Text = @"
Computer Name: $($computer.Name)
DistinguishedName: $($computer.DistinguishedName)
OU: $($computer.DistinguishedName -replace '^CN=.*?,(.*)', '$1')
Last Logon: $([DateTime]::FromFileTime($computer.LastLogonTimestamp))
"@
        $textBox.ReadOnly = $true
        
        $resultForm.Controls.Add($textBox)
        $resultForm.ShowDialog()
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox7.Controls.Add($btnCheckComputer)

$btnMoveComputer = New-Object System.Windows.Forms.Button
$btnMoveComputer.Location = New-Object System.Drawing.Point(350, 110)
$btnMoveComputer.Size = New-Object System.Drawing.Size(150, 30)
$btnMoveComputer.Text = "Move Computer"
$btnMoveComputer.BackColor = "#FF9800"
$btnMoveComputer.ForeColor = "White"
$btnMoveComputer.Add_Click({
    try {
        Get-ADComputer $txtMoveComputer.Text | Move-ADObject -TargetPath $txtTargetOU.Text -Verbose
        [System.Windows.Forms.MessageBox]::Show("Computer moved successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox7.Controls.Add($btnMoveComputer)

$tabPage4.Controls.Add($groupBox7)


# ==== Tab 5: Duplicate Finder ====
$tabPage5 = New-Object System.Windows.Forms.TabPage
$tabPage5.Text = "Duplicate Finder"
$tabPage5.BackColor = "#ffffff"

$groupBox8 = New-Object System.Windows.Forms.GroupBox
$groupBox8.Location = New-Object System.Drawing.Point(20, 20)
$groupBox8.Size = New-Object System.Drawing.Size(800, 540)
$groupBox8.Text = "Search for Duplicate Users"
$groupBox8.ForeColor = "#0066cc"

# Name Label
$labelName = New-Object System.Windows.Forms.Label
$labelName.Location = New-Object System.Drawing.Point(20, 30)
$labelName.Size = New-Object System.Drawing.Size(120, 20)
$labelName.Text = "Full Name:"
$groupBox8.Controls.Add($labelName)

$txtNameInput = New-Object System.Windows.Forms.TextBox
$txtNameInput.Location = New-Object System.Drawing.Point(150, 30)
$txtNameInput.Size = New-Object System.Drawing.Size(250, 25)
$groupBox8.Controls.Add($txtNameInput)

# Email Label
$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Location = New-Object System.Drawing.Point(420, 30)
$labelEmail.Size = New-Object System.Drawing.Size(80, 20)
$labelEmail.Text = "Email:"
$groupBox8.Controls.Add($labelEmail)

$txtEmailInput = New-Object System.Windows.Forms.TextBox
$txtEmailInput.Location = New-Object System.Drawing.Point(500, 30)
$txtEmailInput.Size = New-Object System.Drawing.Size(250, 25)
$groupBox8.Controls.Add($txtEmailInput)

# Strict Mode Checkbox
$chkStrict = New-Object System.Windows.Forms.CheckBox
$chkStrict.Location = New-Object System.Drawing.Point(150, 65)
$chkStrict.Size = New-Object System.Drawing.Size(300, 20)
$chkStrict.Text = "Enable strict first name matching (e.g. 'Luis' ≠ 'Janluis')"
$chkStrict.Checked = $false
$groupBox8.Controls.Add($chkStrict)

# Search Button
$btnSearchDupes = New-Object System.Windows.Forms.Button
$btnSearchDupes.Location = New-Object System.Drawing.Point(500, 65)
$btnSearchDupes.Size = New-Object System.Drawing.Size(150, 30)
$btnSearchDupes.Text = "Find Duplicates"
$btnSearchDupes.BackColor = "#FF9800"
$btnSearchDupes.ForeColor = "White"
$groupBox8.Controls.Add($btnSearchDupes)

# Bulk Upload Button
$btnBulkUpload = New-Object System.Windows.Forms.Button
$btnBulkUpload.Location = New-Object System.Drawing.Point(660, 65)
$btnBulkUpload.Size = New-Object System.Drawing.Size(120, 30)
$btnBulkUpload.Text = "Bulk from CSV"
$btnBulkUpload.BackColor = "#4CAF50"
$btnBulkUpload.ForeColor = "White"
$groupBox8.Controls.Add($btnBulkUpload)

# Results TextBox
$txtResults = New-Object System.Windows.Forms.TextBox
$txtResults.Location = New-Object System.Drawing.Point(20, 110)
$txtResults.Size = New-Object System.Drawing.Size(760, 400)
$txtResults.Multiline = $true
$txtResults.ScrollBars = "Vertical"
$txtResults.ReadOnly = $true
$txtResults.BackColor = "#f9f9f9"
$groupBox8.Controls.Add($txtResults)

# Helper: Search for name/email matches
function Find-Duplicates {
    param(
        [string]$fullName,
        [string]$email,
        [bool]$strict
    )
    $results = @()
    $firstName = ""
    $lastName = "*"

    if ($fullName) {
        $parts = $fullName -split "\s+", 2
        $firstName = $parts[0]
        if ($parts.Length -gt 1) { $lastName = $parts[1] }
    }

    if ($email) {
        $emailMatches = Get-ADUser -Filter { mail -eq $email } -Properties * |
            Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description
        $results += $emailMatches
    }

    if ($firstName) {
        $nameMatches = Get-ADUser -Filter "sn -like '$lastName'" -Properties * |
            Where-Object {
                $gn = $_.GivenName
                if (-not $gn) { return $false }

                if ($strict) {
                    return ($gn -eq $firstName)
                } else {
                    return ($gn -like "*$firstName*") -and
                           ([math]::Abs($gn.Length - $firstName.Length) -le 4)
                }
            } |
            Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description

        $results += $nameMatches

        foreach ($u in $nameMatches) {
            if ($u.mail) {
                $sameMail = Get-ADUser -Filter "mail -eq '$($u.mail)'" -Properties * |
                    Where-Object { $_.SamAccountName -ne $u.SamAccountName } |
                    Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description
                $results += $sameMail
            }
        }
    }

    return $results
}

# Single Search Button
$btnSearchDupes.Add_Click({
    try {
        Import-Module ActiveDirectory
        $name = $txtNameInput.Text.Trim()
        $email = $txtEmailInput.Text.Trim()
        $strict = $chkStrict.Checked

        if (-not $name -and -not $email) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a name and/or email.", "Input Required", "OK", "Warning")
            return
        }

        $results = Find-Duplicates -fullName $name -email $email -strict $strict

        if ($results.Count -eq 0) {
            $txtResults.Text = "No duplicates found."
        } else {
            $distinct = $results | Sort-Object SamAccountName -Unique
            $csv = ($distinct | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
            $txtResults.Text = $csv

            $outPath = "$env:USERPROFILE\Desktop\duplicates_output.csv"
            $distinct | Export-Csv -Path $outPath -NoTypeInformation -Force
            [System.Windows.Forms.MessageBox]::Show("Exported to: $outPath", "Export Complete", "OK", "Info")
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})

# Bulk Upload Button
$btnBulkUpload.Add_Click({
    try {
        Import-Module ActiveDirectory

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv"
        $dialog.Title = "Select CSV with FullName and/or Email headers"

        if ($dialog.ShowDialog() -ne "OK") { return }

        $csv = Import-Csv -Path $dialog.FileName

        $strict = $chkStrict.Checked
        $finalOutput = @()

        foreach ($entry in $csv) {
            $fullName = $entry.FullName
            $email = if ($entry.PSObject.Properties['Email']) { $entry.Email } else { $null }

            if (-not $fullName -and -not $email) { continue }

          if ($email) {
    $header = "# ==== Results for: $fullName [$email] ===="
} else {
    $header = "# ==== Results for: $fullName ===="
}

            $matches = Find-Duplicates -fullName $fullName -email $email -strict $strict

            $finalOutput += $header
            if ($matches.Count -gt 0) {
                $finalOutput += ($matches | Sort-Object SamAccountName -Unique | ConvertTo-Csv -NoTypeInformation)
            } else {
                $finalOutput += "No duplicates found."
            }
            $finalOutput += ""
        }

        $outputPath = "$env:USERPROFILE\Desktop\bulk_duplicates_output.csv"
        $finalOutput -join "`r`n" | Set-Content -Path $outputPath -Encoding UTF8 -Force
        $txtResults.Text = "Bulk duplicate search complete.`r`nOutput saved to:`r`n$outputPath"
        [System.Windows.Forms.MessageBox]::Show("Bulk duplicate check completed.`nSaved to: $outputPath", "Complete", "OK", "Info")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})

$tabPage5.Controls.Add($groupBox8)







# Add all tabs to tab control
$tabControl.Controls.Add($tabPage1)
$tabControl.Controls.Add($tabPage2)
$tabControl.Controls.Add($tabPage3)
$tabControl.Controls.Add($tabPage4)
$tabControl.Controls.Add($tabPage5)

# Add tab control to form
$form.Controls.Add($tabControl)

# Show form
$form.ShowDialog() | Out-Null

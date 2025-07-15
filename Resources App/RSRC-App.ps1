Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "HRA Active Directory Management Tool"
$form.Size = New-Object System.Drawing.Size(850,650)
$form.StartPosition = "CenterScreen"
$form.BackColor = "#f0f0f0"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10,10)
$tabControl.Size = New-Object System.Drawing.Size(810,580)
$tabControl.Anchor = "Top, Bottom, Left, Right"

# Tab 1: Group Management
$tabPage1 = New-Object System.Windows.Forms.TabPage
$tabPage1.Text = "Group Operations"
$tabPage1.BackColor = "#ffffff"

# Group Members Export Section
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Location = New-Object System.Drawing.Point(20,20)
$groupBox1.Size = New-Object System.Drawing.Size(750,150)
$groupBox1.Text = "Export Group Members"
$groupBox1.ForeColor = "#0066cc"

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(20,30)
$label1.Size = New-Object System.Drawing.Size(150,20)
$label1.Text = "Group Name:"
$groupBox1.Controls.Add($label1)

$txtGroupName = New-Object System.Windows.Forms.TextBox
$txtGroupName.Location = New-Object System.Drawing.Point(180,30)
$txtGroupName.Size = New-Object System.Drawing.Size(200,25)
$groupBox1.Controls.Add($txtGroupName)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(20,70)
$label2.Size = New-Object System.Drawing.Size(150,20)
$label2.Text = "Output File Path:"
$groupBox1.Controls.Add($label2)

$txtOutputPath = New-Object System.Windows.Forms.TextBox
$txtOutputPath.Location = New-Object System.Drawing.Point(180,70)
$txtOutputPath.Size = New-Object System.Drawing.Size(400,25)
$txtOutputPath.Text = "C:\temp\group_members.csv"
$groupBox1.Controls.Add($txtOutputPath)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(180,110)
$btnExport.Size = New-Object System.Drawing.Size(150,30)
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

$tabPage1.Controls.Add($groupBox1)

# Group Member Management Section
$groupBox2 = New-Object System.Windows.Forms.GroupBox
$groupBox2.Location = New-Object System.Drawing.Point(20,190)
$groupBox2.Size = New-Object System.Drawing.Size(750,180)
$groupBox2.Text = "Add/Remove Group Members"
$groupBox2.ForeColor = "#0066cc"

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(20,30)
$label3.Size = New-Object System.Drawing.Size(150,20)
$label3.Text = "Group Name:"
$groupBox2.Controls.Add($label3)

$txtTargetGroup = New-Object System.Windows.Forms.TextBox
$txtTargetGroup.Location = New-Object System.Drawing.Point(180,30)
$txtTargetGroup.Size = New-Object System.Drawing.Size(200,25)
$groupBox2.Controls.Add($txtTargetGroup)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(20,70)
$label4.Size = New-Object System.Drawing.Size(150,20)
$label4.Text = "User SAMs (comma sep):"
$groupBox2.Controls.Add($label4)

$txtUserSAMs = New-Object System.Windows.Forms.TextBox
$txtUserSAMs.Location = New-Object System.Drawing.Point(180,70)
$txtUserSAMs.Size = New-Object System.Drawing.Size(400,25)
$txtUserSAMs.Multiline = $true
$txtUserSAMs.Height = 50
$groupBox2.Controls.Add($txtUserSAMs)

$btnAddMembers = New-Object System.Windows.Forms.Button
$btnAddMembers.Location = New-Object System.Drawing.Point(180,130)
$btnAddMembers.Size = New-Object System.Drawing.Size(150,30)
$btnAddMembers.Text = "Add Members"
$btnAddMembers.BackColor = "#4CAF50"
$btnAddMembers.ForeColor = "White"
$btnAddMembers.Add_Click({
    try {
        $members = $txtUserSAMs.Text -split ',' | ForEach-Object { $_.Trim() }
        Add-ADGroupMember -Identity $txtTargetGroup.Text -Members $members
        [System.Windows.Forms.MessageBox]::Show("Users added to group successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox2.Controls.Add($btnAddMembers)

$btnRemoveMembers = New-Object System.Windows.Forms.Button
$btnRemoveMembers.Location = New-Object System.Drawing.Point(350,130)
$btnRemoveMembers.Size = New-Object System.Drawing.Size(150,30)
$btnRemoveMembers.Text = "Remove Members"
$btnRemoveMembers.BackColor = "#f44336"
$btnRemoveMembers.ForeColor = "White"
$btnRemoveMembers.Add_Click({
    try {
        $members = $txtUserSAMs.Text -split ',' | ForEach-Object { $_.Trim() }
        Remove-ADGroupMember -Identity $txtTargetGroup.Text -Members $members -Confirm:$false
        [System.Windows.Forms.MessageBox]::Show("Users removed from group successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox2.Controls.Add($btnRemoveMembers)

$tabPage1.Controls.Add($groupBox2)

# Tab 2: Computer Management
$tabPage2 = New-Object System.Windows.Forms.TabPage
$tabPage2.Text = "Computer Operations"
$tabPage2.BackColor = "#ffffff"

# Computer Reboot Section
$groupBox3 = New-Object System.Windows.Forms.GroupBox
$groupBox3.Location = New-Object System.Drawing.Point(20,20)
$groupBox3.Size = New-Object System.Drawing.Size(750,200)
$groupBox3.Text = "Reboot Computers"
$groupBox3.ForeColor = "#0066cc"

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(20,30)
$label5.Size = New-Object System.Drawing.Size(200,20)
$label5.Text = "Computer Names (one per line):"
$groupBox3.Controls.Add($label5)

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Location = New-Object System.Drawing.Point(20,60)
$txtComputers.Size = New-Object System.Drawing.Size(400,120)
$txtComputers.Multiline = $true
$txtComputers.ScrollBars = "Vertical"
$groupBox3.Controls.Add($txtComputers)

$btnReboot = New-Object System.Windows.Forms.Button
$btnReboot.Location = New-Object System.Drawing.Point(450,60)
$btnReboot.Size = New-Object System.Drawing.Size(150,30)
$btnReboot.Text = "Reboot Computers"
$btnReboot.BackColor = "#FF9800"
$btnReboot.ForeColor = "White"
$btnReboot.Add_Click({
    try {
        $computers = $txtComputers.Text -split "`r`n" | Where-Object { $_ -ne "" }
        foreach ($computer in $computers) {
            Restart-Computer -ComputerName $computer -Force
        }
        [System.Windows.Forms.MessageBox]::Show("Reboot commands sent successfully!", "Success")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox3.Controls.Add($btnReboot)

$tabPage2.Controls.Add($groupBox3)

# Tab 3: User Management
$tabPage3 = New-Object System.Windows.Forms.TabPage
$tabPage3.Text = "User Operations"
$tabPage3.BackColor = "#ffffff"

# User Information Section
$groupBox4 = New-Object System.Windows.Forms.GroupBox
$groupBox4.Location = New-Object System.Drawing.Point(20,20)
$groupBox4.Size = New-Object System.Drawing.Size(750,150)
$groupBox4.Text = "User Information"
$groupBox4.ForeColor = "#0066cc"

$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(20,30)
$label6.Size = New-Object System.Drawing.Size(150,20)
$label6.Text = "Username/SAM:"
$groupBox4.Controls.Add($label6)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(180,30)
$txtUsername.Size = New-Object System.Drawing.Size(200,25)
$groupBox4.Controls.Add($txtUsername)

$btnGetUserInfo = New-Object System.Windows.Forms.Button
$btnGetUserInfo.Location = New-Object System.Drawing.Point(400,30)
$btnGetUserInfo.Size = New-Object System.Drawing.Size(150,30)
$btnGetUserInfo.Text = "Get User Info"
$btnGetUserInfo.BackColor = "#9C27B0"
$btnGetUserInfo.ForeColor = "White"
$btnGetUserInfo.Add_Click({
    try {
        $output = net user $txtUsername.Text /domain
        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "User Information"
        $resultForm.Size = New-Object System.Drawing.Size(500,400)
        
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

$tabPage3.Controls.Add($groupBox4)

# OU User Export Section
$groupBox5 = New-Object System.Windows.Forms.GroupBox
$groupBox5.Location = New-Object System.Drawing.Point(20,190)
$groupBox5.Size = New-Object System.Drawing.Size(750,200)
$groupBox5.Text = "Export Users from OU"
$groupBox5.ForeColor = "#0066cc"

$label7 = New-Object System.Windows.Forms.Label
$label7.Location = New-Object System.Drawing.Point(20,30)
$label7.Size = New-Object System.Drawing.Size(150,20)
$label7.Text = "OU DistinguishedName:"
$groupBox5.Controls.Add($label7)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(180,30)
$txtOU.Size = New-Object System.Drawing.Size(500,25)
$txtOU.Text = "OU=31st Floor,OU=4WTC,OU=People,OU=HRA Resources,DC=windows,DC=nyc,DC=hra,DC=nycnet"
$groupBox5.Controls.Add($txtOU)

$label8 = New-Object System.Windows.Forms.Label
$label8.Location = New-Object System.Drawing.Point(20,70)
$label8.Size = New-Object System.Drawing.Size(150,20)
$label8.Text = "Output File Path:"
$groupBox5.Controls.Add($label8)

$txtOUOutputPath = New-Object System.Windows.Forms.TextBox
$txtOUOutputPath.Location = New-Object System.Drawing.Point(180,70)
$txtOUOutputPath.Size = New-Object System.Drawing.Size(400,25)
$txtOUOutputPath.Text = "C:\temp\ou_users.csv"
$groupBox5.Controls.Add($txtOUOutputPath)

$btnExportOU = New-Object System.Windows.Forms.Button
$btnExportOU.Location = New-Object System.Drawing.Point(180,110)
$btnExportOU.Size = New-Object System.Drawing.Size(150,30)
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

$tabPage3.Controls.Add($groupBox5)

# Add tabs to tab control
$tabControl.Controls.Add($tabPage1)
$tabControl.Controls.Add($tabPage2)
$tabControl.Controls.Add($tabPage3)

# Add tab control to form
$form.Controls.Add($tabControl)

# Show form
$form.ShowDialog() | Out-Null
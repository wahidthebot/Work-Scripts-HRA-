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

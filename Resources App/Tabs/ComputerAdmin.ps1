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
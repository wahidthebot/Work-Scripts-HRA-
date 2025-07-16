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
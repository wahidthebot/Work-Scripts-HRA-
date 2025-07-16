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
        Import-Module ActiveDirectory -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue

        $username = $txtUsername.Text
        if (-not $username) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a username.", "Input Required", "OK", "Information")
            return
        }

        $adUser = Get-ADUser -Identity $username -Properties *
        if (-not $adUser) {
            [System.Windows.Forms.MessageBox]::Show("User not found in Active Directory.", "Error", "OK", "Error")
            return
        }

        $azureUser = $null
        try {
            $azureUser = Get-MgUser -UserId $adUser.UserPrincipalName -Property *
        } catch {
            $azureError = $_.Exception.Message
        }

        $output = @()
        $output += "====== Active Directory Attributes ======"
        $output += $adUser.PSObject.Properties | Sort-Object Name | ForEach-Object {
            "{0,-35}: {1}" -f $_.Name, $_.Value
        }

        if ($azureUser) {
            $output += "`n====== Azure AD Attributes ======"
            $output += $azureUser.PSObject.Properties | Sort-Object Name | ForEach-Object {
                "{0,-35}: {1}" -f $_.Name, $_.Value
            }
        } else {
            $output += "`n====== Azure AD Attributes ======"
            $output += "Azure lookup failed or not connected to Microsoft Graph."
            if ($azureError) {
                $output += "Error: $azureError"
            }
        }

        $resultForm = New-Object System.Windows.Forms.Form
        $resultForm.Text = "User Details"
        $resultForm.Size = New-Object System.Drawing.Size(900, 600)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Multiline = $true
        $textBox.ScrollBars = "Vertical"
        $textBox.Dock = "Fill"
        $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
        $textBox.ReadOnly = $true
        $textBox.Text = $output -join "`r`n"

        $resultForm.Controls.Add($textBox)
        $resultForm.ShowDialog()
    } catch {
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
$txtOUOutputPath.Text = "C:\\temp\\ou_users.csv"
$groupBox5.Controls.Add($txtOUOutputPath)

$btnExportOU = New-Object System.Windows.Forms.Button
$btnExportOU.Location = New-Object System.Drawing.Point(180, 110)
$btnExportOU.Size = New-Object System.Drawing.Size(150, 30)
$btnExportOU.Text = "Export Users"
$btnExportOU.BackColor = "#0066cc"
$btnExportOU.ForeColor = "White"
$btnExportOU.Add_Click({
    try {
        Get-ADUser -Filter * -SearchBase $txtOU.Text -Properties extensionAttribute2 | 
            Select-Object SamAccountName, extensionAttribute2 | 
            Export-Csv -Path $txtOUOutputPath.Text -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("Users exported successfully!", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})
$groupBox5.Controls.Add($btnExportOU)

$tabPage3.Controls.Add($groupBox4)
$tabPage3.Controls.Add($groupBox5)

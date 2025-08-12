# ===== TAB 3: USER MANAGEMENT =====
$tabPage3 = New-Object System.Windows.Forms.TabPage
$tabPage3.Text = "User Management"
$tabPage3.BackColor = "#ffffff"

# User Information Group
$groupBox4 = New-Object System.Windows.Forms.GroupBox
$groupBox4.Location = New-Object System.Drawing.Point(20, 20)
$groupBox4.Size = New-Object System.Drawing.Size(1200, 600)
$groupBox4.Text = "User Information"
$groupBox4.ForeColor = "#0066cc"

# Username Label
$label6 = New-Object System.Windows.Forms.Label
$label6.Location = New-Object System.Drawing.Point(20, 30)
$label6.Size = New-Object System.Drawing.Size(150, 20)
$label6.Text = "Username/SAM:"
$groupBox4.Controls.Add($label6)

# Username Textbox
$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(180, 30)
$txtUsername.Size = New-Object System.Drawing.Size(200, 25)
$groupBox4.Controls.Add($txtUsername)

# Get User Info Button
$btnGetUserInfo = New-Object System.Windows.Forms.Button
$btnGetUserInfo.Location = New-Object System.Drawing.Point(400, 30)
$btnGetUserInfo.Size = New-Object System.Drawing.Size(150, 30)
$btnGetUserInfo.Text = "Get User Info"
$btnGetUserInfo.BackColor = "#9C27B0"
$btnGetUserInfo.ForeColor = "White"
$groupBox4.Controls.Add($btnGetUserInfo)

# Search Box
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(20, 70)
$searchBox.Size = New-Object System.Drawing.Size(300, 25)
$searchBox.ForeColor = "Gray"
$searchBox.Text = "Search attributes..."
$searchBox.Add_GotFocus({
    if ($searchBox.ForeColor -eq "Gray") {
        $searchBox.Text = ""
        $searchBox.ForeColor = "Black"
    }
})
$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.ForeColor = "Gray"
        $searchBox.Text = "Search attributes..."
    }
})
$groupBox4.Controls.Add($searchBox)

# AD Attributes CheckedListBox
$chkListAD = New-Object System.Windows.Forms.CheckedListBox
$chkListAD.Location = New-Object System.Drawing.Point(20, 110)
$chkListAD.Size = New-Object System.Drawing.Size(500, 400)
$chkListAD.CheckOnClick = $true
$groupBox4.Controls.Add($chkListAD)

# Azure Attributes CheckedListBox
$chkListAzure = New-Object System.Windows.Forms.CheckedListBox
$chkListAzure.Location = New-Object System.Drawing.Point(540, 110)
$chkListAzure.Size = New-Object System.Drawing.Size(500, 400)
$chkListAzure.CheckOnClick = $true
$groupBox4.Controls.Add($chkListAzure)

# Output Textbox
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Location = New-Object System.Drawing.Point(20, 520)
$outputBox.Size = New-Object System.Drawing.Size(1020, 60)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$groupBox4.Controls.Add($outputBox)

# Show Selected Button
$btnShowSelected = New-Object System.Windows.Forms.Button
$btnShowSelected.Text = "Show Selected"
$btnShowSelected.Location = New-Object System.Drawing.Point(340, 70)
$btnShowSelected.Size = New-Object System.Drawing.Size(120, 30)
$groupBox4.Controls.Add($btnShowSelected)

# Variables to store props for filtering
$script:adProps = @()
$script:azureProps = @()

# Get User Info Click Event
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
        $azureError = $null
        try {
            $azureUser = Get-MgUser -UserId $adUser.UserPrincipalName -Property *
        } catch {
            $azureError = $_.Exception.Message
        }

        # Store properties for filtering
        $script:adProps = $adUser.PSObject.Properties | Sort-Object Name
        $script:azureProps = if ($azureUser) { $azureUser.PSObject.Properties | Sort-Object Name } else { @() }

        # Populate AD list
        $chkListAD.Items.Clear()
        foreach ($prop in $script:adProps) {
            $chkListAD.Items.Add("$($prop.Name): $($prop.Value)")
        }

        # Populate Azure list
        $chkListAzure.Items.Clear()
        if ($azureUser) {
            foreach ($prop in $script:azureProps) {
                $chkListAzure.Items.Add("$($prop.Name): $($prop.Value)")
            }
        } else {
            $chkListAzure.Items.Add("Azure lookup failed or not connected to Microsoft Graph.")
            if ($azureError) { $chkListAzure.Items.Add("Error: $azureError") }
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})

# Search filtering
$searchBox.Add_TextChanged({
    if ($searchBox.ForeColor -eq "Gray") { return }
    $searchText = $searchBox.Text.ToLower()

    # Filter AD
    $chkListAD.Items.Clear()
    foreach ($prop in $script:adProps) {
        if ($prop.Name.ToLower() -like "*$searchText*" -or "$($prop.Value)".ToLower() -like "*$searchText*") {
            $chkListAD.Items.Add("$($prop.Name): $($prop.Value)")
        }
    }

    # Filter Azure
    $chkListAzure.Items.Clear()
    foreach ($prop in $script:azureProps) {
        if ($prop.Name.ToLower() -like "*$searchText*" -or "$($prop.Value)".ToLower() -like "*$searchText*") {
            $chkListAzure.Items.Add("$($prop.Name): $($prop.Value)")
        }
    }
})

# Show Selected Click Event
$btnShowSelected.Add_Click({
    $selected = @()
    $selected += $chkListAD.CheckedItems
    $selected += $chkListAzure.CheckedItems
    $outputBox.Text = $selected -join "`r`n"
})

# Add to tab
$tabPage3.Controls.Add($groupBox4)

# ================= Export Users from OU ==================
$groupBox5 = New-Object System.Windows.Forms.GroupBox
$groupBox5.Location = New-Object System.Drawing.Point(20, 640)
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

$tabPage3.Controls.Add($groupBox5)

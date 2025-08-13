# ==== Tab 5: Duplicate Finder ====
$tabPage5 = New-Object System.Windows.Forms.TabPage
$tabPage5.Text = "Duplicate Finder"
$tabPage5.BackColor = "#ffffff"

$groupBox8 = New-Object System.Windows.Forms.GroupBox
$groupBox8.Location = New-Object System.Drawing.Point(20, 20)
$groupBox8.Size = New-Object System.Drawing.Size(1600, 620)
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
$txtNameInput.Size = New-Object System.Drawing.Size(500, 25)
$groupBox8.Controls.Add($txtNameInput)

# Email Label
$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Location = New-Object System.Drawing.Point(670, 30)
$labelEmail.Size = New-Object System.Drawing.Size(80, 20)
$labelEmail.Text = "Email:"
$groupBox8.Controls.Add($labelEmail)

$txtEmailInput = New-Object System.Windows.Forms.TextBox
$txtEmailInput.Location = New-Object System.Drawing.Point(760, 30)
$txtEmailInput.Size = New-Object System.Drawing.Size(500, 25)
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
$btnSearchDupes.Location = New-Object System.Drawing.Point(760, 65)
$btnSearchDupes.Size = New-Object System.Drawing.Size(150, 30)
$btnSearchDupes.Text = "Find Duplicates"
$btnSearchDupes.BackColor = "#FF9800"
$btnSearchDupes.ForeColor = "White"
$groupBox8.Controls.Add($btnSearchDupes)

# Bulk Upload Button
$btnBulkUpload = New-Object System.Windows.Forms.Button
$btnBulkUpload.Location = New-Object System.Drawing.Point(920, 65)
$btnBulkUpload.Size = New-Object System.Drawing.Size(120, 30)
$btnBulkUpload.Text = "Bulk from CSV"
$btnBulkUpload.BackColor = "#4CAF50"
$btnBulkUpload.ForeColor = "White"
$groupBox8.Controls.Add($btnBulkUpload)

# Results TextBox
$txtResults = New-Object System.Windows.Forms.TextBox
$txtResults.Location = New-Object System.Drawing.Point(20, 110)
$txtResults.Size = New-Object System.Drawing.Size(1560, 380)
$txtResults.Multiline = $true
$txtResults.ScrollBars = "Vertical"
$txtResults.ReadOnly = $true
$txtResults.BackColor = "#f9f9f9"
$groupBox8.Controls.Add($txtResults)

# ==== Column Selection ====
$labelColumns = New-Object System.Windows.Forms.Label
$labelColumns.Location = New-Object System.Drawing.Point(20, 500)
$labelColumns.Size = New-Object System.Drawing.Size(250, 20)
$labelColumns.Text = "Select columns to include in CSV:"
$groupBox8.Controls.Add($labelColumns)

function Add-ColumnCheckbox {
    param([string]$text, [int]$x, [int]$y)
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $text
    $cb.Checked = $true
    $cb.Location = New-Object System.Drawing.Point($x, $y)
    $cb.Size = New-Object System.Drawing.Size(170, 20)
    $groupBox8.Controls.Add($cb)
    return $cb
}

$cbGivenName       = Add-ColumnCheckbox "GivenName"         20  520
$cbSurname         = Add-ColumnCheckbox "Surname"           210 520
$cbSamAccountName  = Add-ColumnCheckbox "SamAccountName"    380 520
$cbMail            = Add-ColumnCheckbox "mail"              590 520
$cbUPN             = Add-ColumnCheckbox "UserPrincipalName" 760 520
$cbEnabled         = Add-ColumnCheckbox "Enabled"           930 520
$cbDescription     = Add-ColumnCheckbox "Description"        980 520
$cbDistinguished   = Add-ColumnCheckbox "DistinguishedName" 1155 520  # NEW FIELD

function Get-SelectedColumns {
    $cols = @()
    if ($cbGivenName.Checked)      { $cols += "GivenName" }
    if ($cbSurname.Checked)        { $cols += "Surname" }
    if ($cbSamAccountName.Checked) { $cols += "SamAccountName" }
    if ($cbMail.Checked)           { $cols += "mail" }
    if ($cbUPN.Checked)            { $cols += "UserPrincipalName" }
    if ($cbEnabled.Checked)        { $cols += "Enabled" }
    if ($cbDescription.Checked)    { $cols += "Description" }
    if ($cbDistinguished.Checked)  { $cols += "DistinguishedName" }  # NEW FIELD
    return $cols
}

# Helper: Search for name/email matches
function Find-Duplicates {
    param(
        [string]$fullName,
        [string]$email,
        [bool]$strict
    )

    $results = @()

    # Handle email first
    if ($email) {
        $emailMatches = Get-ADUser -Filter "mail -eq '$email'" -Properties * |
            Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description, DistinguishedName
        $results += $emailMatches
    }

    if ($fullName) {
        # Split name into parts
        $nameParts = $fullName -split '\s+' | Where-Object { $_ -ne '' }

        # Build LDAP filter
        $filterTerms = foreach ($part in $nameParts) {
            "(|" +
                "(givenName=*$part*)" +
                "(sn=*$part*)" +
                "(displayName=*$part*)" +
                "(cn=*$part*)" +
                "(samAccountName=*$part*)" +
                "(userPrincipalName=*$part*)" +
                "(mail=*$part*)" +
            ")"
        }

        $finalFilter = "(&" + ($filterTerms -join "") + ")"

        $nameMatches = Get-ADUser -LDAPFilter $finalFilter -Properties GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description, DistinguishedName |
            Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description, DistinguishedName

        $results += $nameMatches

        foreach ($user in $nameMatches) {
            if ($user.mail) {
                $dupes = Get-ADUser -Filter "mail -eq '$($user.mail)'" -Properties * |
                    Where-Object { $_.SamAccountName -ne $user.SamAccountName } |
                    Select GivenName, Surname, SamAccountName, mail, UserPrincipalName, Enabled, Description, DistinguishedName
                $results += $dupes
            }
        }
    }

    return $results
}

# Single Search Button
$btnSearchDupes.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
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
            $cols = Get-SelectedColumns
            $distinct = $results | Sort-Object SamAccountName -Unique | Select-Object $cols
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
        Import-Module ActiveDirectory -ErrorAction Stop

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "CSV files (*.csv)|*.csv"
        $dialog.Title = "Select CSV with FullName and/or Email headers"

        if ($dialog.ShowDialog() -ne "OK") { return }

        $csv = Import-Csv -Path $dialog.FileName

        $strict = $chkStrict.Checked
        $finalOutput = @()
        $cols = Get-SelectedColumns

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
                $finalOutput += ($matches | Sort-Object SamAccountName -Unique | Select-Object $cols | ConvertTo-Csv -NoTypeInformation)
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

# Group Extraction handlers
$extractButton.Add_Click({
    $groupName = $groupNameTextBox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($groupName)) {
        $outputTextBox.AppendText("Error: Group name is required.`n")
        return
    }
    
    try {
        $outputTextBox.AppendText("Starting group extraction...`n")
        $progressBar.Value = 10
        $statusLabel.Content = "Connecting to Active Directory..."
        
        # Check if group exists
        $group = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction Stop
        if (-not $group) {
            $outputTextBox.AppendText("Error: Group '$groupName' not found in Active Directory.`n")
            $progressBar.Value = 0
            $statusLabel.Content = "Ready"
            return
        }
        
        $progressBar.Value = 20
        $statusLabel.Content = "Retrieving group members..."
        $outputTextBox.AppendText("Group found. Retrieving members...`n")
        
        # Get all members of the group
        $members = Get-ADGroupMember -Identity $groupName -Recursive | 
                   Where-Object { $_.objectClass -eq 'user' } | 
                   Get-ADUser -Properties *
        
        $progressBar.Value = 40
        $statusLabel.Content = "Processing user data..."
        $outputTextBox.AppendText("Found $($members.Count) users. Processing data...`n")
        
        # Extract user information
        $script:extractedData = @()
        $userCount = $members.Count
        $currentUser = 0
        
        foreach ($user in $members) {
            $currentUser++
            $progress = 40 + ([math]::Round(($currentUser / $userCount) * 50))
            $progressBar.Value = $progress
            $statusLabel.Content = "Processing user $currentUser of $userCount..."
            
            $userInfo = [PSCustomObject]@{
                Name = $user.Name
                SamAccountName = $user.SamAccountName
                UserPrincipalName = $user.UserPrincipalName
                Email = $user.EmailAddress
                Title = $user.Title
                Department = $user.Department
                Company = $user.Company
                Office = $user.Office
                StreetAddress = $user.StreetAddress
                City = $user.City
                State = $user.State
                PostalCode = $user.PostalCode
                Country = $user.Country
                Telephone = $user.telephoneNumber
                Mobile = $user.mobile
                EmployeeID = $user.EmployeeID
                EmployeeType = $user.EmployeeType
                Enabled = $user.Enabled
                LastLogonDate = $user.LastLogonDate
                PasswordLastSet = $user.PasswordLastSet
                AccountExpirationDate = $user.AccountExpirationDate
                WhenCreated = $user.WhenCreated
                DistinguishedName = $user.DistinguishedName
            }
            
            $script:extractedData += $userInfo
        }
        
        $progressBar.Value = 100
        $statusLabel.Content = "Extraction complete"
        $outputTextBox.AppendText("Successfully extracted data for $($script:extractedData.Count) users.`n")
        $downloadButton.IsEnabled = $true
        
    } catch {
        $outputTextBox.AppendText("Error during extraction: $_`n")
        $progressBar.Value = 0
        $statusLabel.Content = "Error occurred"
    }
})

$downloadButton.Add_Click({
    if (-not $script:extractedData -or $script:extractedData.Count -eq 0) {
        $outputTextBox.AppendText("Error: No data to export. Please extract data first.`n")
        return
    }
    
    try {
        $progressBar.Value = 0
        $statusLabel.Content = "Preparing export..."
        $outputTextBox.AppendText("Preparing Excel export...`n")
        
        # Create a SaveFileDialog
        $saveFileDialog = New-Object Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $saveFileDialog.FileName = "$($groupNameTextBox.Text)_UserExport_$(Get-Date -Format 'yyyyMMdd').xlsx"
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            $script:exportFilePath = $saveFileDialog.FileName
            $outputTextBox.AppendText("Exporting to: $script:exportFilePath`n")
            
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            
            # Create Excel objects
            $progressBar.Value = 10
            $statusLabel.Content = "Creating Excel file..."
            
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "User Data"
            
            # Add headers
            $progressBar.Value = 20
            $statusLabel.Content = "Adding headers..."
            $column = 1
            $script:extractedData[0].PSObject.Properties.Name | ForEach-Object {
                $worksheet.Cells.Item(1, $column) = $_
                $worksheet.Cells.Item(1, $column).Font.Bold = $true
                $worksheet.Cells.Item(1, $column).Interior.ColorIndex = 15
                $column++
            }
            
            # Add data
            $progressBar.Value = 30
            $statusLabel.Content = "Adding data..."
            $row = 2
            $totalRows = $script:extractedData.Count
            $currentRow = 0
            
            foreach ($user in $script:extractedData) {
                $currentRow++
                $progress = 30 + ([math]::Round(($currentRow / $totalRows) * 60))
                $progressBar.Value = $progress
                $statusLabel.Content = "Exporting row $currentRow of $totalRows..."
                
                $column = 1
                $user.PSObject.Properties.Value | ForEach-Object {
                    $worksheet.Cells.Item($row, $column) = $_
                    $column++
                }
                $row++
            }
            
            # Auto-fit columns
            $progressBar.Value = 95
            $statusLabel.Content = "Formatting document..."
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
            
            # Add freeze panes and formatting
            $worksheet.Activate()
            $worksheet.Application.ActiveWindow.SplitRow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true
            
            # Save and close
            $progressBar.Value = 98
            $statusLabel.Content = "Saving file..."
            $workbook.SaveAs($script:exportFilePath)
            $workbook.Close($false)
            
            $progressBar.Value = 100
            $statusLabel.Content = "Export complete"
            $outputTextBox.AppendText("Successfully exported data to Excel file.`n")
            
            # Offer to open the file
            $result = [System.Windows.MessageBox]::Show("Export completed successfully. Would you like to open the file now?", "Export Complete", "YesNo", "Question")
            if ($result -eq "Yes") {
                Start-Process $script:exportFilePath
            }
        }
    } catch {
        $outputTextBox.AppendText("Error during export: $_`n")
        $progressBar.Value = 0
        $statusLabel.Content = "Error occurred"
    } finally {
        if ($workbook) { $workbook.Close($false) }
        $progressBar.Value = 0
        $statusLabel.Content = "Ready"
    }
})

try {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
} catch {
    Write-Warning "Error cleaning up Excel COM objects: $_"
}

# Main Launcher Script - PROPERLY FIXED VERSION
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get the directory of the currently running script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the relative path to the XAML file
$xamlFilePath = Join-Path -Path $scriptDir -ChildPath "MainWindow.xaml"

# Check if XAML file exists
if (-not (Test-Path $xamlFilePath)) {
    [System.Windows.Forms.MessageBox]::Show("XAML file not found at: $xamlFilePath", "Error", "OK", "Error")
    return
}

# Load the XAML
[xml]$xaml = Get-Content -Path $xamlFilePath
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# ===== FIND ALL CONTROLS FIRST =====
# Main menu buttons
$remoteAccessButton = $window.FindName("RemoteAccessButton")
$lanExtensionButton = $window.FindName("LanExtensionButton")
$groupManagementButton = $window.FindName("GroupManagementButton")
$MFAUpdateButton = $window.FindName("MFAUpdateButton")
$groupExtractionButton = $window.FindName("GroupExtractionButton")
$createUserButton = $window.FindName("CreateUserButton")
$adResetB2BButton = $window.FindName("AdResetB2BButton")

# Panels
$remoteAccessPanel = $window.FindName("RemoteAccessPanel")
$lanExtensionPanel = $window.FindName("LanExtensionPanel")
$groupManagementPanel = $window.FindName("GroupManagementPanel")
$MFAUpdatePanel = $window.FindName("MFAUpdatePanel")
$groupExtractionPanel = $window.FindName("GroupExtractionPanel")
$createUserPanel = $window.FindName("CreateUserPanel")
$adResetB2BPanel = $window.FindName("AdResetB2BPanel")

# Output textbox - THIS IS CRITICAL
$outputTextBox = $window.FindName("OutputTextBox")

# Remote Access controls
$computerNameTextBox = $window.FindName("ComputerNameTextBox")
$userLANIDTextBox = $window.FindName("UserLANIDTextBox")
$groupSelectionComboBox = $window.FindName("GroupSelectionComboBox")
$remoteAccessSubmitButton = $window.FindName("RemoteAccessSubmitButton")

# LAN Extension controls
$lanExtensionUserLANIDTextBox = $window.FindName("LanExtensionUserLANIDTextBox")
$lanExtensionDateTextBox = $window.FindName("LanExtensionDateTextBox")
$lanExtensionTicketNumberTextBox = $window.FindName("LanExtensionTicketNumberTextBox")
$lanExtensionInitialsTextBox = $window.FindName("LanExtensionInitialsTextBox")
$lanExtensionSubmitButton = $window.FindName("LanExtensionSubmitButton")

# Group Management controls
$groupManagementUserLANIDTextBox = $window.FindName("GroupManagementUserLANIDTextBox")
$groupManagementGroupNamesTextBox = $window.FindName("GroupManagementGroupNamesTextBox")
$groupManagementActionComboBox = $window.FindName("GroupManagementActionComboBox")
$groupManagementSubmitButton = $window.FindName("GroupManagementSubmitButton")

# MFA Update controls
$MFAUpdateUserEmailTextBox = $window.FindName("MFAUpdateUserEmailTextBox")
$MFAUpdatePhoneNumberTextBox = $window.FindName("MFAUpdatePhoneNumberTextBox")
$MFAUpdateMethodTypeComboBox = $window.FindName("MFAUpdateMethodTypeComboBox")
$MFAUpdateSubmitButton = $window.FindName("MFAUpdateSubmitButton")

# Group Extraction controls
$groupNameTextBox = $window.FindName("GroupNameTextBox")
$extractButton = $window.FindName("ExtractButton")
$downloadButton = $window.FindName("DownloadButton")
$progressBar = $window.FindName("ProgressBar")
$statusLabel = $window.FindName("StatusLabel")

# Create User controls
$createUserSubmitButton = $window.FindName("CreateUserSubmitButton")
$createUserLoadCSVButton = $window.FindName("CreateUserLoadCSVButton")
$createUserClearButton = $window.FindName("CreateUserClearButton")
$createUserFirstNameTextBox = $window.FindName("CreateUserFirstNameTextBox")
$createUserLastNameTextBox = $window.FindName("CreateUserLastNameTextBox")
$createUserTitleTextBox = $window.FindName("CreateUserTitleTextBox")
$createUserDepartmentTextBox = $window.FindName("CreateUserDepartmentTextBox")
$createUserPhoneTextBox = $window.FindName("CreateUserPhoneTextBox")
$createUserEmailTextBox = $window.FindName("CreateUserEmailTextBox")
$createUserDescriptionTextBox = $window.FindName("CreateUserDescriptionTextBox")
$createUserGroupsTextBox = $window.FindName("CreateUserGroupsTextBox")
$createUserOUTextBox = $window.FindName("CreateUserOUTextBox")
$createUserLanIDTextBox = $window.FindName("CreateUserLanIDTextBox")
$createUserOrgComboBox = $window.FindName("CreateUserOrgComboBox")
$createUserTypeComboBox = $window.FindName("CreateUserTypeComboBox")
$createUserExpiryLabel = $window.FindName("CreateUserExpiryLabel")
$createUserExpiryComboBox = $window.FindName("CreateUserExpiryComboBox")

# AD Reset & B2B controls
$adResetB2BUserTextBox = $window.FindName("AdResetB2BUserTextBox")
$adResetPasswordButton = $window.FindName("AdResetPasswordButton")
$adResetGuestButton = $window.FindName("AdResetGuestButton")
$adResetMemberButton = $window.FindName("AdResetMemberButton")

# ===== PANEL VISIBILITY HANDLERS =====
$remoteAccessButton.Add_Click({
    $remoteAccessPanel.Visibility = "Visible"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Remote Access panel selected.`n")
})

$lanExtensionButton.Add_Click({
    $lanExtensionPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("LAN Extension panel selected.`n")
})

$groupManagementButton.Add_Click({
    $groupManagementPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Group Management panel selected.`n")
})

$MFAUpdateButton.Add_Click({
    $MFAUpdatePanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("MFA Update panel selected.`n")
})

$groupExtractionButton.Add_Click({
    $groupExtractionPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Group Extraction panel selected.`n")
})

$createUserButton.Add_Click({
    $createUserPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $adResetB2BPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("Create User panel selected.`n")
})

$adResetB2BButton.Add_Click({
    $adResetB2BPanel.Visibility = "Visible"
    $remoteAccessPanel.Visibility = "Collapsed"
    $lanExtensionPanel.Visibility = "Collapsed"
    $groupManagementPanel.Visibility = "Collapsed"
    $MFAUpdatePanel.Visibility = "Collapsed"
    $groupExtractionPanel.Visibility = "Collapsed"
    $createUserPanel.Visibility = "Collapsed"
    $outputTextBox.Clear()
    $outputTextBox.AppendText("AD Reset & B2B panel selected.`n")
})

# ===== CREATE GLOBAL VARIABLES FOR FUNCTION FILES =====
# These variables will be available to the function files
$global:outputTextBox = $outputTextBox
$global:window = $window

# ===== LOAD FUNCTION MODULES =====
# Define function file paths
$functionFiles = @(
    "$scriptDir\Functions\RemoteAccessFunctions.ps1",
    "$scriptDir\Functions\LanExtensionFunctions.ps1", 
    "$scriptDir\Functions\GroupManagementFunctions.ps1",
    "$scriptDir\Functions\MFAUpdateFunctions.ps1",
    "$scriptDir\Functions\GroupExtractionFunctions.ps1",
    "$scriptDir\Functions\CreateUserFunctions.ps1",
    "$scriptDir\Functions\AdResetB2BFunctions.ps1"
)

# Load each function file if it exists
foreach ($file in $functionFiles) {
    if (Test-Path $file) {
        try {
            . $file
            $outputTextBox.AppendText("Loaded: $(Split-Path $file -Leaf)`n")
        } catch {
            $outputTextBox.AppendText("Error loading $(Split-Path $file -Leaf): $_`n")
        }
    } else {
        $outputTextBox.AppendText("Warning: File not found: $(Split-Path $file -Leaf)`n")
    }
}

# ===== ATTACH EVENT HANDLERS FOR FILES THAT MIGHT HAVE FAILED =====
# These are fallbacks in case the function files didn't attach handlers properly

# Remote Access fallback
if ($remoteAccessSubmitButton -and ($remoteAccessSubmitButton.HasEvents -eq $false -or -not $remoteAccessSubmitButton.GetType().GetEvent("Click").GetAddMethod())) {
    $remoteAccessSubmitButton.Add_Click({
        $outputTextBox.AppendText("Remote Access function would execute here (fallback).`n")
        $outputTextBox.AppendText("Computer: $($computerNameTextBox.Text), User: $($userLANIDTextBox.Text), Group: $($groupSelectionComboBox.SelectedItem.Content)`n")
    })
}

# LAN Extension fallback
if ($lanExtensionSubmitButton -and ($lanExtensionSubmitButton.HasEvents -eq $false -or -not $lanExtensionSubmitButton.GetType().GetEvent("Click").GetAddMethod())) {
    $lanExtensionSubmitButton.Add_Click({
        $outputTextBox.AppendText("LAN Extension function would execute here (fallback).`n")
    })
}

# Group Management fallback
if ($groupManagementSubmitButton -and ($groupManagementSubmitButton.HasEvents -eq $false -or -not $groupManagementSubmitButton.GetType().GetEvent("Click").GetAddMethod())) {
    $groupManagementSubmitButton.Add_Click({
        $outputTextBox.AppendText("Group Management function would execute here (fallback).`n")
    })
}

# MFA Update fallback
if ($MFAUpdateSubmitButton -and ($MFAUpdateSubmitButton.HasEvents -eq $false -or -not $MFAUpdateSubmitButton.GetType().GetEvent("Click").GetAddMethod())) {
    $MFAUpdateSubmitButton.Add_Click({
        $outputTextBox.AppendText("MFA Update function would execute here (fallback).`n")
    })
}

# Create User Type ComboBox fallback
if ($createUserTypeComboBox -and ($createUserTypeComboBox.HasEvents -eq $false -or -not $createUserTypeComboBox.GetType().GetEvent("SelectionChanged").GetAddMethod())) {
    $createUserTypeComboBox.Add_SelectionChanged({
        if ($createUserTypeComboBox.SelectedItem) {
            $isTemp = $createUserTypeComboBox.SelectedItem.Content -match "Temp"
            $createUserExpiryLabel.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
            $createUserExpiryComboBox.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
        }
    })
}

# ===== INITIAL SETUP =====
# Initialize variables for group extraction
$script:extractedData = $null
$script:exportFilePath = $null

# Show initial message
$outputTextBox.AppendText("Account Management App initialized.`n")
$outputTextBox.AppendText("Select a function from the menu above.`n")

# ===== SHOW WINDOW =====
$window.ShowDialog() | Out-Null
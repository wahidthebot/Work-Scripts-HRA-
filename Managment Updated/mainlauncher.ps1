# Main Launcher Script - ULTIMATE FIXED VERSION
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

# ===== CREATE GLOBAL VARIABLES FOR COMBOBOX SELECTIONS =====
# Initialize global variables for ComboBox selections FIRST
$global:SelectedAccountType = "Default"
$global:SelectedPCFarmGroup = "HRARDPUsers1"

# ===== FIND ALL CONTROLS FIRST =====
# Use try-catch for each control to prevent errors if control not found
try {
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
    $AccountTypeComboBox = $window.FindName("AccountTypeComboBox")
    $PCFarmComboBox = $window.FindName("PCFarmComboBox")
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
    $createUserOUComboBox = $window.FindName("CreateUserOUComboBox")
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

    Write-Host "All controls loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error loading controls: $_" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("Error loading UI controls: $_", "Error", "OK", "Error")
    return
}

# ===== REGISTER COMBOBOX EVENT HANDLERS =====
if ($AccountTypeComboBox -ne $null) {
    $AccountTypeComboBox.Add_SelectionChanged({
        if ($AccountTypeComboBox.SelectedItem -ne $null) {
            $global:SelectedAccountType = $AccountTypeComboBox.SelectedItem.Content.ToString().Trim()
            Write-Host "Account Type set to: $global:SelectedAccountType" -ForegroundColor Yellow
        }
    })
}

if ($PCFarmComboBox -ne $null) {
    $PCFarmComboBox.Add_SelectionChanged({
        if ($PCFarmComboBox.SelectedItem -ne $null) {
            $global:SelectedPCFarmGroup = $PCFarmComboBox.SelectedItem.Content.ToString().Trim()
            Write-Host "PC Farm Group set to: $global:SelectedPCFarmGroup" -ForegroundColor Yellow
        }
    })
}

# ===== PANEL VISIBILITY HANDLERS =====
if ($remoteAccessButton -ne $null) {
    $remoteAccessButton.Add_Click({
        $remoteAccessPanel.Visibility = "Visible"
        $lanExtensionPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("Remote Access panel selected.`n")
        }
    })
}

if ($lanExtensionButton -ne $null) {
    $lanExtensionButton.Add_Click({
        $lanExtensionPanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("LAN Extension panel selected.`n")
        }
    })
}

if ($groupManagementButton -ne $null) {
    $groupManagementButton.Add_Click({
        $groupManagementPanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $lanExtensionPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("Group Management panel selected.`n")
        }
    })
}

if ($MFAUpdateButton -ne $null) {
    $MFAUpdateButton.Add_Click({
        $MFAUpdatePanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $lanExtensionPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("MFA Update panel selected.`n")
        }
    })
}

if ($groupExtractionButton -ne $null) {
    $groupExtractionButton.Add_Click({
        $groupExtractionPanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $lanExtensionPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("Group Extraction panel selected.`n")
        }
    })
}

if ($createUserButton -ne $null) {
    $createUserButton.Add_Click({
        $createUserPanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $lanExtensionPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $adResetB2BPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("Create User panel selected.`n")
        }
    })
}

if ($adResetB2BButton -ne $null) {
    $adResetB2BButton.Add_Click({
        $adResetB2BPanel.Visibility = "Visible"
        $remoteAccessPanel.Visibility = "Collapsed"
        $lanExtensionPanel.Visibility = "Collapsed"
        $groupManagementPanel.Visibility = "Collapsed"
        $MFAUpdatePanel.Visibility = "Collapsed"
        $groupExtractionPanel.Visibility = "Collapsed"
        $createUserPanel.Visibility = "Collapsed"
        if ($outputTextBox -ne $null) {
            $outputTextBox.Clear()
            $outputTextBox.AppendText("AD Reset & B2B panel selected.`n")
        }
    })
}

# ===== CREATE GLOBAL VARIABLES FOR FUNCTION FILES =====
# These variables will be available to the function files
$global:outputTextBox = $outputTextBox
$global:window = $window

# Also export the specific Remote Access controls to global scope for the function file
$global:ComputerNameTextBox = $computerNameTextBox
$global:UserLANIDTextBox = $userLANIDTextBox
$global:AccountTypeComboBox = $AccountTypeComboBox
$global:PCFarmComboBox = $PCFarmComboBox
$global:RemoteAccessSubmitButton = $remoteAccessSubmitButton

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

# Create Functions directory if it doesn't exist
$functionsDir = Join-Path -Path $scriptDir -ChildPath "Functions"
if (-not (Test-Path $functionsDir)) {
    New-Item -ItemType Directory -Path $functionsDir -Force | Out-Null
    Write-Host "Created Functions directory" -ForegroundColor Yellow
}

# Load each function file if it exists
foreach ($file in $functionFiles) {
    if (Test-Path $file) {
        try {
            . $file
            Write-Host "Loaded: $(Split-Path $file -Leaf)" -ForegroundColor Green
            if ($outputTextBox -ne $null) {
                $outputTextBox.AppendText("Loaded: $(Split-Path $file -Leaf)`n")
            }
        } catch {
            Write-Host "Error loading $(Split-Path $file -Leaf): $_" -ForegroundColor Red
            if ($outputTextBox -ne $null) {
                $outputTextBox.AppendText("Error loading $(Split-Path $file -Leaf): $_`n")
            }
        }
    } else {
        Write-Host "Warning: File not found: $(Split-Path $file -Leaf)" -ForegroundColor Yellow
        if ($outputTextBox -ne $null) {
            $outputTextBox.AppendText("Warning: File not found: $(Split-Path $file -Leaf)`n")
        }
    }
}

# ===== ATTACH EVENT HANDLERS FOR FILES THAT MIGHT HAVE FAILED =====
# These are fallbacks in case the function files didn't attach handlers properly

# Create User Type ComboBox fallback for expiry visibility
if ($createUserTypeComboBox -ne $null) {
    $createUserTypeComboBox.Add_SelectionChanged({
        if ($createUserTypeComboBox.SelectedItem) {
            $isTemp = $createUserTypeComboBox.SelectedItem.Content -match "Temp"
            if ($createUserExpiryLabel -ne $null) {
                $createUserExpiryLabel.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
            }
            if ($createUserExpiryComboBox -ne $null) {
                $createUserExpiryComboBox.Visibility = if ($isTemp) { "Visible" } else { "Collapsed" }
            }
        }
    })
}

# Remote Access fallback - ONLY IF NOT ALREADY HANDLED
if ($remoteAccessSubmitButton -ne $null) {
    $remoteAccessSubmitButton.Add_Click({
        if ($outputTextBox -ne $null) {
            $outputTextBox.AppendText("Remote Access function would execute here (fallback).`n")
            $outputTextBox.AppendText("Computer: $($computerNameTextBox.Text), User: $($userLANIDTextBox.Text)`n")
            $outputTextBox.AppendText("Account Type: $global:SelectedAccountType, PC Farm: $global:SelectedPCFarmGroup`n")
        }
    })
}


# ===== INITIAL SETUP =====
# Initialize variables for group extraction
$script:extractedData = $null
$script:exportFilePath = $null

# Show initial message
if ($outputTextBox -ne $null) {
    $outputTextBox.AppendText("Account Management App initialized.`n")
    $outputTextBox.AppendText("Select a function from the menu above.`n")
}

# Add Window.Closing event to prevent issues
$window.Add_Closing({
    param($sender, $e)
    Write-Host "Window is closing..." -ForegroundColor Yellow
})

# ===== SHOW WINDOW =====
try {
    Write-Host "Showing window..." -ForegroundColor Green
    $window.ShowDialog() | Out-Null
    Write-Host "Window closed successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error showing window: $_" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("Error showing window: $_", "Error", "OK", "Error")
}

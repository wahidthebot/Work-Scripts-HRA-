Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "HRA Active Directory Management Tool"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = "#f0f0f0"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(860, 630)
$tabControl.Anchor = "Top, Bottom, Left, Right"

# Load each tab from external scripts
. "$PSScriptRoot\Tabs\GroupManagement.ps1"
. "$PSScriptRoot\Tabs\ComputerManagement.ps1"
. "$PSScriptRoot\Tabs\UserManagement.ps1"
. "$PSScriptRoot\Tabs\ComputerAdmin.ps1"
. "$PSScriptRoot\Tabs\DuplicateFinder.ps1"

# Add tabs to the tab control (assuming each script sets $tabPage1, $tabPage2, etc.)
$tabControl.Controls.Add($tabPage1)
$tabControl.Controls.Add($tabPage2)
$tabControl.Controls.Add($tabPage3)
$tabControl.Controls.Add($tabPage4)
$tabControl.Controls.Add($tabPage5)

# Add tab control to form
$form.Controls.Add($tabControl)

# Show form
$form.ShowDialog()

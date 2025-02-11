# RevitCloudOpener.ps1

# Script Information
$script:SCRIPT_VERSION = "2.0.1"
$script:LAST_UPDATED = "2025-02-05 13:59:23"
$script:CURRENT_USER = "HKR (Moin)"

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Constants
$script:MODELS_CSV_PATH = "L:\HKR BIM\BIM360_Models.csv"
$script:BATCH_RVT_PATH = "$env:LOCALAPPDATA\RevitBatchProcessor\BatchRvt.exe"
$script:TEMP_FOLDER = "$env:TEMP\RevitCloudOpener"
$script:REVIT_SCRIPT_HOST_PATH = "$env:LOCALAPPDATA\RevitBatchProcessor\Scripts\revit_script_host.py"

# Create temp folder if it doesn't exist
if (-not (Test-Path $script:TEMP_FOLDER)) {
    New-Item -ItemType Directory -Path $script:TEMP_FOLDER | Out-Null
}

# Default script content
$script:DEFAULT_SCRIPT_CONTENT = @'
import revit_script_util
from revit_script_util import Output
doc = revit_script_util.GetScriptDocument()
'@

# Create the minimal Python script
$taskScriptPath = Join-Path $script:TEMP_FOLDER "minimal_script.dyn"
Set-Content -Path $taskScriptPath -Value $script:DEFAULT_SCRIPT_CONTENT

function Update-ScriptHostFile {
    param (
        [bool]$enableFastMode
    )
    
    try {
        if (Test-Path $script:REVIT_SCRIPT_HOST_PATH) {
            $content = Get-Content $script:REVIT_SCRIPT_HOST_PATH -Raw
            
            if ($enableFastMode) {
                $content = $content -replace "currentProcess\.Kill\(\)", "currentProcess.Refresh()"
                $content = $content -replace "process\.CloseMainWindow\(\)", "process.Refresh()"
                $content = $content -replace "END_SESSION_DELAY_IN_SECONDS = 5", "END_SESSION_DELAY_IN_SECONDS = 1"
            } else {
                $content = $content -replace "currentProcess\.Refresh\(\)", "currentProcess.Kill()"
                $content = $content -replace "process\.Refresh\(\)", "process.CloseMainWindow()"
                $content = $content -replace "END_SESSION_DELAY_IN_SECONDS = 1", "END_SESSION_DELAY_IN_SECONDS = 5"
            }
            
            Set-Content -Path $script:REVIT_SCRIPT_HOST_PATH -Value $content
            return $true
        } else {
            [System.Windows.Forms.MessageBox]::Show("revit_script_host.py file not found", "Error")
            return $false
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error updating script host file", "Error")
        return $false
    }
}

function Check-ScriptHostMode {
    try {
        if (Test-Path $script:REVIT_SCRIPT_HOST_PATH) {
            $content = Get-Content $script:REVIT_SCRIPT_HOST_PATH -Raw
            return $content -match "currentProcess\.Refresh\(\)" -and 
                   $content -match "process\.Refresh\(\)" -and 
                   $content -match "END_SESSION_DELAY_IN_SECONDS = 1"
        }
        return $false
    }
    catch {
        return $false
    }
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Revit Cloud Model Opener - v$script:SCRIPT_VERSION"
$form.Size = New-Object System.Drawing.Size(1200,800)
$form.StartPosition = "CenterScreen"

# Create control panel
$controlPanel = New-Object System.Windows.Forms.Panel
$controlPanel.Location = New-Object System.Drawing.Point(10,10)
$controlPanel.Size = New-Object System.Drawing.Size(1160,180)
$form.Controls.Add($controlPanel)

# Create options group box
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Text = "Model Opening Options"
$optionsGroup.Location = New-Object System.Drawing.Point(10,10)
$optionsGroup.Size = New-Object System.Drawing.Size(320,100)
$controlPanel.Controls.Add($optionsGroup)

# Create radio buttons for model options
$createNewLocal = New-Object System.Windows.Forms.RadioButton
$createNewLocal.Text = "Create New Local"
$createNewLocal.Location = New-Object System.Drawing.Point(10,25)
$createNewLocal.Size = New-Object System.Drawing.Size(140,20)
$createNewLocal.Checked = $true
$optionsGroup.Controls.Add($createNewLocal)

$detachCentral = New-Object System.Windows.Forms.RadioButton
$detachCentral.Text = "Detach from Central"
$detachCentral.Location = New-Object System.Drawing.Point(150,25)
$detachCentral.Size = New-Object System.Drawing.Size(140,20)
$optionsGroup.Controls.Add($detachCentral)

# Create checkboxes
$worksetCheckbox = New-Object System.Windows.Forms.CheckBox
$worksetCheckbox.Text = "Load Worksets"
$worksetCheckbox.Location = New-Object System.Drawing.Point(10,55)
$worksetCheckbox.Size = New-Object System.Drawing.Size(140,20)
$worksetCheckbox.Checked = $true
$optionsGroup.Controls.Add($worksetCheckbox)

$auditCheckbox = New-Object System.Windows.Forms.CheckBox
$auditCheckbox.Text = "Perform Audit"
$auditCheckbox.Location = New-Object System.Drawing.Point(150,55)
$auditCheckbox.Size = New-Object System.Drawing.Size(140,20)
$optionsGroup.Controls.Add($auditCheckbox)

# Create script selection group
$scriptGroup = New-Object System.Windows.Forms.GroupBox
$scriptGroup.Text = "Script Selection"
$scriptGroup.Location = New-Object System.Drawing.Point(10,110)
$scriptGroup.Size = New-Object System.Drawing.Size(1130,70)
$controlPanel.Controls.Add($scriptGroup)

# Create radio buttons for script selection
$defaultScript = New-Object System.Windows.Forms.RadioButton
$defaultScript.Text = "No Script"
$defaultScript.Location = New-Object System.Drawing.Point(10,30)
$defaultScript.Size = New-Object System.Drawing.Size(120,20)
$defaultScript.Checked = $true
$scriptGroup.Controls.Add($defaultScript)

$customScript = New-Object System.Windows.Forms.RadioButton
$customScript.Text = "Custom Script"
$customScript.Location = New-Object System.Drawing.Point(150,30)
$customScript.Size = New-Object System.Drawing.Size(150,20)
$scriptGroup.Controls.Add($customScript)

# Add browse button for custom script
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse..."
$browseButton.Location = New-Object System.Drawing.Point(300,28)
$browseButton.Size = New-Object System.Drawing.Size(100,24)
$browseButton.Enabled = $false
$scriptGroup.Controls.Add($browseButton)

# Add script path label
$scriptPathLabel = New-Object System.Windows.Forms.Label
$scriptPathLabel.Text = "No custom script selected"
$scriptPathLabel.Location = New-Object System.Drawing.Point(450,28)
$scriptPathLabel.Size = New-Object System.Drawing.Size(600,20)
$scriptPathLabel.AutoEllipsis = $true
$scriptGroup.Controls.Add($scriptPathLabel)

# Create filter group
$filterGroup = New-Object System.Windows.Forms.GroupBox
$filterGroup.Text = "Filters"
$filterGroup.Location = New-Object System.Drawing.Point(340,10)
$filterGroup.Size = New-Object System.Drawing.Size(800,100)
$controlPanel.Controls.Add($filterGroup)

# Project filter label
$projectFilterLabel = New-Object System.Windows.Forms.Label
$projectFilterLabel.Text = "Project:"
$projectFilterLabel.Location = New-Object System.Drawing.Point(10,20)
$projectFilterLabel.Size = New-Object System.Drawing.Size(60,20)
$filterGroup.Controls.Add($projectFilterLabel)

# Project filter combo box
$projectFilter = New-Object System.Windows.Forms.ComboBox
$projectFilter.Location = New-Object System.Drawing.Point(70,20)
$projectFilter.Size = New-Object System.Drawing.Size(250,30)
$projectFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$filterGroup.Controls.Add($projectFilter)

# Path filter label
$pathFilterLabel = New-Object System.Windows.Forms.Label
$pathFilterLabel.Text = "Filter Path:"
$pathFilterLabel.Location = New-Object System.Drawing.Point(350,20)
$pathFilterLabel.Size = New-Object System.Drawing.Size(100,20)
$filterGroup.Controls.Add($pathFilterLabel)

# Path filter textbox
$pathFilter = New-Object System.Windows.Forms.TextBox
$pathFilter.Location = New-Object System.Drawing.Point(450,20)
$pathFilter.Size = New-Object System.Drawing.Size(180,20)
$filterGroup.Controls.Add($pathFilter)

# File name filter label
$fileNameFilterLabel = New-Object System.Windows.Forms.Label
$fileNameFilterLabel.Text = "Filter File Name:"
$fileNameFilterLabel.Location = New-Object System.Drawing.Point(350,60)
$fileNameFilterLabel.Size = New-Object System.Drawing.Size(100,20)
$filterGroup.Controls.Add($fileNameFilterLabel)

# File name filter textbox
$fileNameFilter = New-Object System.Windows.Forms.TextBox
$fileNameFilter.Location = New-Object System.Drawing.Point(450,60)
$fileNameFilter.Size = New-Object System.Drawing.Size(180,20)
$filterGroup.Controls.Add($fileNameFilter)

# Add Font Size Control
$fontSizeGroup = New-Object System.Windows.Forms.GroupBox
$fontSizeGroup.Text = "Font Size"
$fontSizeGroup.Location = New-Object System.Drawing.Point(650,50)
$fontSizeGroup.Size = New-Object System.Drawing.Size(150,80)
$filterGroup.Controls.Add($fontSizeGroup)

$fontSizeNumeric = New-Object System.Windows.Forms.NumericUpDown
$fontSizeNumeric.Location = New-Object System.Drawing.Point(50,20)
$fontSizeNumeric.Size = New-Object System.Drawing.Size(50,50)
$fontSizeNumeric.Minimum = 8
$fontSizeNumeric.Maximum = 20
$fontSizeNumeric.Value = 11
$fontSizeNumeric.TextAlign = "Center"
$fontSizeGroup.Controls.Add($fontSizeNumeric)

# Refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(20,60)
$refreshButton.Size = New-Object System.Drawing.Size(140,30)
$refreshButton.Text = "Refresh Models"
$filterGroup.Controls.Add($refreshButton)

# Create Fast Mode toggle button
$scriptHostButton = New-Object System.Windows.Forms.Button
$scriptHostButton.Location = New-Object System.Drawing.Point(170,60)
$scriptHostButton.Size = New-Object System.Drawing.Size(150,30)
$scriptHostButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

# Check initial state
$isFastMode = Check-ScriptHostMode
if ($isFastMode) {
    $scriptHostButton.Text = "Keep Revit open (ON)"
    $scriptHostButton.BackColor = [System.Drawing.Color]::LightGreen
} else {
    $scriptHostButton.Text = "Keep Revit open (OFF)"
    $scriptHostButton.BackColor = [System.Drawing.Color]::LightCoral
}
$filterGroup.Controls.Add($scriptHostButton)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,200)
$listView.Size = New-Object System.Drawing.Size(1160,490)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.MultiSelect = $false

# Add columns (modified)
$listView.Columns.Add("File Name", 580)
$listView.Columns.Add("Folder Path", 580)

$form.Controls.Add($listView)

# Create status strip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready Credits : HKR (Moin)"
$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# Function to load models
function Load-Models {
    param(
        $selectedProject = "All Projects",
        $pathFilter = "",
        $fileNameFilter = ""
    )
    
    try {
        $listView.Items.Clear()
        
        if (Test-Path $script:MODELS_CSV_PATH) {
            $models = Import-Csv $script:MODELS_CSV_PATH
            
            if ($selectedProject -ne "All Projects") {
                $models = $models | Where-Object { $_.'Project Name' -eq $selectedProject }
            }
            
            if ($pathFilter) {
                $models = $models | Where-Object { 
                    $_.'Folder Path' -like "*$pathFilter*" -or
                    $_.'Source File Name' -like "*$pathFilter*"
                }
            }

            if ($fileNameFilter) {
                $models = $models | Where-Object { 
                    $_.'Source File Name' -like "*$fileNameFilter*"
                }
            }
            
            foreach ($model in $models) {
                $item = New-Object System.Windows.Forms.ListViewItem($model.'Source File Name')
                $item.SubItems.Add($model.'Folder Path')
                
                $item.Tag = @{
                    ProjectGuid = $model.'Project GUID'
                    ModelGuid = $model.'Model GUID'
                    FileName = $model.'Source File Name'
                    RevitVersion = $model.'Revit Version'
                }
                
                $listView.Items.Add($item)
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading models", "Error")
    }
}

# Function to populate project filter
function Initialize-ProjectFilter {
    try {
        if (Test-Path $script:MODELS_CSV_PATH) {
            $models = Import-Csv $script:MODELS_CSV_PATH
            $projectFilter.Items.Clear()
            $projectFilter.Items.Add("All Projects")
            $models | Select-Object -ExpandProperty 'Project Name' -Unique | Sort-Object | ForEach-Object {
                $projectFilter.Items.Add($_)
            }
            $projectFilter.SelectedIndex = 0
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading projects", "Error")
    }
}

# Find and update the Open-RevitModel function
function Open-RevitModel {
    param($modelData)
    
    try {
        $tempFileList = Join-Path $script:TEMP_FOLDER "model_list.txt"
        $modelLine = "$($modelData.RevitVersion) $($modelData.ProjectGuid) $($modelData.ModelGuid)"
        Set-Content -Path $tempFileList -Value $modelLine
        
        # Determine which script path to use
        $scriptToUse = if ($customScript.Checked -and $scriptPathLabel.Text -ne "No custom script selected") {
            # Use the full path from the custom script selection
            $scriptPathLabel.Text
        } else {
            # Use the default minimal script
            $taskScriptPath
        }
        
        # Build arguments based on options
        $arguments = "--file_list `"$tempFileList`" --task_script `"$scriptToUse`""
        
        if ($createNewLocal.Checked) {
            $arguments += " --create_new_local"
        } else {
            $arguments += " --detach"
        }
        
        if ($worksetCheckbox.Checked) {
            $arguments += " --worksets open_all"
        } else {
            $arguments += " --worksets close_all"
        }
        
        if ($auditCheckbox.Checked) {
            $arguments += " --audit"
        }
        
        # Start BatchRvt process with visible window
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $script:BATCH_RVT_PATH
        $processInfo.Arguments = $arguments
        $processInfo.UseShellExecute = $true
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
        
        $process = [System.Diagnostics.Process]::Start($processInfo)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error opening model", "Error")
    }
}

# Event handlers
$refreshButton.Add_Click({
    Load-Models $projectFilter.SelectedItem $pathFilter.Text $fileNameFilter.Text
})

$projectFilter.Add_SelectedIndexChanged({
    Load-Models $projectFilter.SelectedItem $pathFilter.Text $fileNameFilter.Text
})

$pathFilter.Add_TextChanged({
    Load-Models $projectFilter.SelectedItem $pathFilter.Text $fileNameFilter.Text
})

$fileNameFilter.Add_TextChanged({
    Load-Models $projectFilter.SelectedItem $pathFilter.Text $fileNameFilter.Text
})

# Font size change handler
$fontSizeNumeric.Add_ValueChanged({
    $newSize = $fontSizeNumeric.Value
    $listView.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $newSize)
})

$scriptHostButton.Add_Click({
    $currentState = Check-ScriptHostMode
    $success = Update-ScriptHostFile -enableFastMode (!$currentState)
    
    if ($success) {
        if (!$currentState) {
            $scriptHostButton.Text = "Keep Revit open (ON)"
            $scriptHostButton.BackColor = [System.Drawing.Color]::LightGreen
        } else {
            $scriptHostButton.Text = "Keep Revit open (OFF)"
            $scriptHostButton.BackColor = [System.Drawing.Color]::LightCoral
        }
    }
})

# Add event handlers for script selection
$customScript.Add_CheckedChanged({
    $browseButton.Enabled = $customScript.Checked
})

$defaultScript.Add_CheckedChanged({
    if ($defaultScript.Checked) {
        $scriptPathLabel.Text = "No custom script selected"
        Set-Content -Path $taskScriptPath -Value $script:DEFAULT_SCRIPT_CONTENT
    }
})

$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Dynamo Script (*.dyn)|*.dyn|All files (*.*)|*.*"
    $openFileDialog.Title = "Select Custom Script"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $customScriptContent = Get-Content -Path $openFileDialog.FileName -Raw
            $scriptPathLabel.Text = $openFileDialog.FileName
            Set-Content -Path $taskScriptPath -Value $customScriptContent
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading custom script", "Error")
            $defaultScript.Checked = $true
            $scriptPathLabel.Text = "No custom script selected"
        }
    }
})

$listView.Add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        Open-RevitModel $selectedItem.Tag
    }
})

# Form cleanup
$form.Add_FormClosing({
    param($sender, $e)
    
    # Clean up temporary files
    if (Test-Path $script:TEMP_FOLDER) {
        Remove-Item -Path $script:TEMP_FOLDER -Recurse -Force -ErrorAction SilentlyContinue
    }
})

# Initialize and show form
Initialize-ProjectFilter
Load-Models
# Set initial font size
$listView.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $fontSizeNumeric.Value)
$form.ShowDialog()
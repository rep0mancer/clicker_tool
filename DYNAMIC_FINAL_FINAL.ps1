#region Assembly Loading
# Load necessary .NET assemblies for GUI elements
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion Assembly Loading

#region MouseSimulator Class
# Defines a class to interact with the low-level mouse_event API
$mouseSignature = @"
using System;
using System.Runtime.InteropServices;
public class MouseSimulator {
    // Import the mouse_event function from user32.dll
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

    // Define constants for mouse actions for clarity and readability
    public const uint MOUSEEVENTF_LEFTDOWN = 0x0002; // Left mouse button down
    public const uint MOUSEEVENTF_LEFTUP = 0x0004;   // Left mouse button up
    public const uint MOUSEEVENTF_MOVE = 0x0001;     // Mouse move
    public const uint MOUSEEVENTF_ABSOLUTE = 0x8000; // Absolute mouse coordinates
}
"@

# Check if the "MouseSimulator" type already exists in the current application domain.
# This prevents errors if the script is run multiple times in the same PowerShell session.
$existingType = [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {
    $_.GetType("MouseSimulator", $false) # Attempt to get the type without throwing an error
} | Where-Object { $_ -ne $null } # Filter out null results (type not found in assembly)
if (-not $existingType) {
    # If the type doesn't exist, add it using the C# code defined above.
    Add-Type -TypeDefinition $mouseSignature -PassThru | Out-Null
}

# Function to simulate a left mouse button click (down then up)
function Click-MouseButton {
    [MouseSimulator]::mouse_event([MouseSimulator]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    [MouseSimulator]::mouse_event([MouseSimulator]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
}

# Function to move the mouse cursor to a relative position on the screen and then click
function Move-And-Click {
    param(
        [Parameter(Mandatory=$true)]
        [double]$xRel, # Relative X coordinate (0.0 to 1.0)
        [Parameter(Mandatory=$true)]
        [double]$yRel  # Relative Y coordinate (0.0 to 1.0)
    )
    # Get the bounds of the primary screen
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    # Basic validation for screen dimensions
    if ($screen.Width -le 0 -or $screen.Height -le 0) {
        Write-Error "Invalid screen dimensions detected ($($screen.Width)x$($screen.Height)). Cannot calculate cursor position."
        return
    }
    # Calculate absolute X and Y coordinates from relative inputs
    $x = [int]($screen.Width * $xRel)
    $y = [int]($screen.Height * $yRel)
    # Set the cursor position
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
    # Short delay to allow UI to potentially update before clicking
    Start-Sleep -Milliseconds 100 # Consider making this configurable or removing if unnecessary
    # Perform the click
    Click-MouseButton
}
#endregion MouseSimulator Class

#region Global Parameters and Actions
# Define global delay settings used during automation playback
$global:DelayBeforeClick   = 500 # Milliseconds delay after moving mouse, before clicking
$global:DelayAfterClick    = 500 # Milliseconds delay after sending input via SendKeys
$global:DelayBetweenActions = 1500 # Milliseconds delay after completing each action (Click or Sleep)

# Initialize the global list to store the sequence of actions (using ArrayList for flexibility)
$global:Actions = [System.Collections.ArrayList]::new()

# Define regex patterns for identifying variable placeholders in action inputs
$global:VariableInputPlaceholderPattern = '^%%VARIABLE_INPUT_(\d+)%%$' # For exact match (less used now)
$global:VariableInputScanPattern = '%%VARIABLE_INPUT_(\d+)%%' # For finding placeholders anywhere within a string
#endregion Global Parameters and Actions

#region GUI Functions and Elements

# Updates the listbox displaying the recorded action sequence
function Update-ListBox {
    Write-Host "[DEBUG] Updating ListBox..."
    $listBox.BeginUpdate() # Prevent flickering during updates
    $listBox.Items.Clear() # Remove all existing items
    # Iterate through the global actions list
    for ($i = 0; $i -lt $global:Actions.Count; $i++) {
        $action = $global:Actions[$i] # Get the current action object
        # Format the display text based on the action type
        $displayText = "Action $([int]($i+1)): $($action.Type)" # Start with action number and type
        if ($action.Type -eq "Click") {
            $displayText += " at ($($action.XRel), $($action.YRel))" # Add coordinates for Click actions
            # Add associated input text if it exists and is not empty
            if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                 $displayText += " - Input: $($action.Input)"
            }
        } elseif ($action.Type -eq "Sleep") {
            $displayText += " for $($action.Duration) ms" # Add duration for Sleep actions
        }
        $listBox.Items.Add($displayText) | Out-Null # Add the formatted string to the listbox
    }
    $listBox.EndUpdate() # Re-enable listbox drawing
    Write-Host "[DEBUG] ListBox Updated with $($global:Actions.Count) items."
}

# Helper function to add a new InputN column to the DataGridView if it doesn't already exist
function Add-GridInputColumn {
    param(
        [Parameter(Mandatory=$true)]
        [ref]$dataGridViewRef, # Pass DataGridView by reference to modify the original object
        [Parameter(Mandatory=$true)]
        [int]$columnNumberToAdd # The number N for the "InputN" column
    )
    $dataGridView = $dataGridViewRef.Value # Dereference to get the actual DataGridView object
    $colName = "Input$columnNumberToAdd" # Construct the internal column name
    $headerText = "Input $columnNumberToAdd" # Construct the user-visible header text
    # Check if a column with this name already exists
    if (-not $dataGridView.Columns.Contains($colName)) {
        # Create a new text box column
        $newCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $newCol.Name = $colName         # Set the internal name
        $newCol.HeaderText = $headerText # Set the display header
        # Optional: Set minimum width or other properties
        # $newCol.MinimumWidth = 60
        $dataGridView.Columns.Add($newCol) | Out-Null # Add the column to the grid
        Write-Host "[DEBUG] Helper Function: Added grid column: $headerText" # Log the addition
        return $true # Indicate that the column was successfully added
    }
    # Return false if the column already existed
    return $false
}

# Factory function for creating Label controls
function Create-Label { param([int]$x,[int]$y,[int]$width,[int]$height,[string]$text) $l=New-Object System.Windows.Forms.Label; $l.Location=New-Object System.Drawing.Point($x,$y); $l.Size=New-Object System.Drawing.Size($width,$height); $l.Text=$text; return $l }

# --- MODIFIED Create-Button Function with Debugging ---
# Factory function for creating Button controls
function Create-Button {
    param(
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [string]$text,
        [scriptblock]$clickAction # The script block to execute on click
    )
    Write-Host "[DEBUG] Creating Button: '$text'" # DEBUG Start
    $b = New-Object System.Windows.Forms.Button
    $b.Location = New-Object System.Drawing.Point($x,$y)
    $b.Size = New-Object System.Drawing.Size($width,$height)
    $b.Text = $text
    # Check if a click action script block was provided
    if ($clickAction -ne $null) {
        Write-Host "[DEBUG]   Attempting to attach Click Action for '$text'" # DEBUG Before Add_Click
        try {
            # Attach the provided script block to the button's Click event
            $b.Add_Click($clickAction)
            Write-Host "[DEBUG]   Click Action attached successfully for '$text'." # DEBUG After Add_Click success
        } catch {
            # Log any error during the attachment process (should be rare)
             Write-Host "[DEBUG]   ERROR attaching Click Action for '$text': $($_.Exception.Message)" # DEBUG Error
        }
    } else {
         # Log if no click action was provided (useful for debugging missing actions)
         Write-Host "[DEBUG]   WARNING: No Click Action provided for '$text'." # DEBUG No action provided
    }
    return $b # Return the created button object
}
# --- END OF MODIFIED Create-Button Function ---

# Factory function for creating TextBox controls
function Create-TextBox { param([int]$x,[int]$y,[int]$width,[int]$height) $tb=New-Object System.Windows.Forms.TextBox; $tb.Location=New-Object System.Drawing.Point($x,$y); $tb.Size=New-Object System.Drawing.Size($width,$height); return $tb }
# Factory function for creating TrackBar controls
function Create-TrackBar { param([int]$x,[int]$y,[int]$width,[int]$min,[int]$max,[int]$tick,[int]$val,[scriptblock]$changed) $tk=New-Object System.Windows.Forms.TrackBar;$tk.Location=New-Object System.Drawing.Point($x,$y);$tk.Width=$width;$tk.Minimum=$min;$tk.Maximum=$max;$tk.TickFrequency=$tick;$tk.Value=$val;$tk.Add_ValueChanged($changed);return $tk }

#endregion GUI Functions and Elements


#region Form Creation and Controls
# Create the main application window (Form)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Steiner's Automatisierer QoL (Debug 2)" # Window title updated
$form.Size = New-Object System.Drawing.Size(750, 900) # Initial size (Width, Height)
$form.StartPosition = "CenterScreen" # Start the form in the center of the screen
$form.KeyPreview = $true # Allow the form to receive key events before controls do (for F9/Escape)

# --- Top Instruction Label ---
$instructionLabel = Create-Label -x 10 -y 10 -width ($form.ClientSize.Width - 20) -height 30 -text "F9: Capture. Esc: Cancel Edit. Define sequence, use %%VARIABLE_INPUT_N%%, manage grid data."
$instructionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right # Anchor to top, left, right
$form.Controls.Add($instructionLabel)

# --- Action Sequence ListBox --- (Left Side)
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,50)
$listBox.Size = New-Object System.Drawing.Size(400,150) # Define size
$listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right # Anchor relative to form edges
$form.Controls.Add($listBox)

# --- Action List Controls --- (Right Side)
$actionButtonX = 420 # Starting X coordinate for the right-side button group
$actionButtonWidth = 100 # Standard width for these buttons
$actionButtonGap = 5   # Horizontal gap between button columns

# Column 1 of action buttons (Copy, Clear All)
$copyButton = Create-Button -x $actionButtonX -y 50 -width $actionButtonWidth -height 30 -text "Copy Action" -clickAction {
    Write-Host "[EVENT] Copy Action Button Clicked" # DEBUG Event Trigger
    $selectedIndex = $listBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        Write-Host "[EVENT]   Item selected at index: $selectedIndex" # DEBUG
        $originalAction = $global:Actions[$selectedIndex]
        # Create a deep copy using JSON roundtrip to ensure independence
        Write-Host "[EVENT]   Original Action: $($originalAction | ConvertTo-Json -Depth 2 -Compress)" # DEBUG
        $newAction = ($originalAction | ConvertTo-Json -Depth 5 | ConvertFrom-Json)
        # Insert the copy immediately after the original
        $global:Actions.Insert($selectedIndex + 1, $newAction)
        Write-Host "[EVENT]   Action copied and inserted at index $($selectedIndex + 1)" # DEBUG
        Update-ListBox # Refresh the listbox display
        $listBox.SelectedIndex = $selectedIndex + 1 # Select the newly created copy
    } else {
        Write-Host "[EVENT]   No item selected to copy" # DEBUG
        # Inform user if no action is selected
        [System.Windows.Forms.MessageBox]::Show("Please select an action in the list to copy.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}
$copyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right # Anchor button
$form.Controls.Add($copyButton)

$clearButton = Create-Button -x $actionButtonX -y 90 -width $actionButtonWidth -height 30 -text "Clear All" -clickAction {
    Write-Host "[EVENT] Clear All Button Clicked" # DEBUG Event Trigger
    # Confirm before clearing all actions
    if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear all actions in the sequence?", "Confirm Clear", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq 'Yes') {
        Write-Host "[EVENT]   Clearing all actions..." # DEBUG
        $global:Actions.Clear() # Clear the global actions list
        Update-ListBox # Refresh the display
        Write-Host "[EVENT]   Actions cleared." # DEBUG
    } else {
        Write-Host "[EVENT]   Clear cancelled by user." # DEBUG
    }
}
$clearButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($clearButton)

# Column 2 of action buttons (Move Up, Move Down, Delete, Edit)
$col2X = $actionButtonX + $actionButtonWidth + $actionButtonGap # Calculate X for the second column
$moveUpButton = Create-Button -x $col2X -y 50 -width $actionButtonWidth -height 30 -text "Move Up" -clickAction {
    # Expanded logic for clarity
    Write-Host "[EVENT] Move Up Button Clicked" # DEBUG Event Trigger
    $index = $listBox.SelectedIndex
    if ($index -gt 0) { # Check if an item is selected and it's not the first one
        Write-Host "[EVENT]   Moving item at index $index up" # DEBUG
        $itemToMove = $global:Actions[$index] # Get the selected item
        $global:Actions.RemoveAt($index) # Remove it from its current position
        $global:Actions.Insert($index - 1, $itemToMove) # Insert it one position higher
        Update-ListBox # Refresh the display
        $listBox.SelectedIndex = $index - 1 # Select the item in its new position
    } else {
        Write-Host "[EVENT]   Cannot move item up (not selected or already at top)" # DEBUG
    }
}
$moveUpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($moveUpButton)

$moveDownButton = Create-Button -x $col2X -y 90 -width $actionButtonWidth -height 30 -text "Move Down" -clickAction {
    # Expanded logic for clarity
    Write-Host "[EVENT] Move Down Button Clicked" # DEBUG Event Trigger
    $index = $listBox.SelectedIndex
    # Check if an item is selected and it's not the last one
    if ($index -ge 0 -and $index -lt ($global:Actions.Count - 1)) {
        Write-Host "[EVENT]   Moving item at index $index down" # DEBUG
        $itemToMove = $global:Actions[$index] # Get the selected item
        $global:Actions.RemoveAt($index) # Remove it from its current position
        $global:Actions.Insert($index + 1, $itemToMove) # Insert it one position lower
        Update-ListBox # Refresh the display
        $listBox.SelectedIndex = $index + 1 # Select the item in its new position
    } else {
         Write-Host "[EVENT]   Cannot move item down (not selected or already at bottom)" # DEBUG
    }
}
$moveDownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($moveDownButton)

$deleteButton = Create-Button -x $col2X -y 130 -width $actionButtonWidth -height 30 -text "Delete" -clickAction {
    Write-Host "[EVENT] Delete Button Clicked" # DEBUG Event Trigger
    $index = $listBox.SelectedIndex
    if ($index -ge 0) { # Check if an item is selected
        Write-Host "[EVENT]   Deleting item at index $index" # DEBUG
        $global:Actions.RemoveAt($index) # Remove the item
        Update-ListBox # Refresh the display
    } else {
        Write-Host "[EVENT]   No item selected to delete" # DEBUG
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list to delete.")
    }
}
$deleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($deleteButton)

$editActionButton = Create-Button -x $col2X -y 170 -width $actionButtonWidth -height 30 -text "Edit Action" -clickAction {
    # Expanded logic for clarity
    Write-Host "[EVENT] Edit Action Button Clicked" # DEBUG Event Trigger
    $idx = $listBox.SelectedIndex;
    if ($idx -ge 0) {
        Write-Host "[EVENT]   Editing item at index $idx" # DEBUG
        $selectedAction = $global:Actions[$idx];
        # Enable the editing groupbox and its controls
        $editActionGroupBox.Enabled = $true
        $updateActionButton.Enabled = $true
        $editCancelButton.Enabled = $true
        $editActionTypeTextBox.Text = $selectedAction.Type; # Display the action type (read-only)
        Write-Host "[EVENT]   Action Type: $($selectedAction.Type)" # DEBUG
        # Show/hide controls based on action type and populate with current values
        if ($selectedAction.Type -eq "Click") {
             Write-Host "[EVENT]   Populating Click fields..." # DEBUG
             $editXRelLabel.Visible = $true
             $editXRelTextBox.Visible = $true
             $editYRelLabel.Visible = $true
             $editYRelTextBox.Visible = $true
             $editInputLabel.Visible = $true
             $editInputTextBox.Visible = $true
             $editDurationLabel.Visible = $false
             $editDurationTextBox.Visible = $false; # Hide sleep controls
             # Populate click controls
             $editXRelTextBox.Text = $selectedAction.XRel
             $editYRelTextBox.Text = $selectedAction.YRel
             $editInputTextBox.Text = $selectedAction.Input
        } elseif ($selectedAction.Type -eq "Sleep") {
             Write-Host "[EVENT]   Populating Sleep fields..." # DEBUG
             $editDurationLabel.Visible = $true
             $editDurationTextBox.Visible = $true; # Show sleep controls
             $editXRelLabel.Visible = $false
             $editXRelTextBox.Visible = $false
             $editYRelLabel.Visible = $false
             $editYRelTextBox.Visible = $false
             $editInputLabel.Visible = $false
             $editInputTextBox.Visible = $false # Hide click controls
             # Populate sleep control
             $editDurationTextBox.Text = $selectedAction.Duration
        }
        Write-Host "[EVENT]   Edit GroupBox enabled and populated." # DEBUG
    } else {
        Write-Host "[EVENT]   No item selected to edit." # DEBUG
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list to edit.")
    }
}
$editActionButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($editActionButton)

# --- Insert Sleep --- (Right Side, below action buttons)
$sleepY = 210 # Y position for sleep controls relative to top
$sleepLabel = Create-Label -x $actionButtonX -y $sleepY -width 150 -height 20 -text "Sleep Duration (ms):"
$sleepLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sleepLabel)
$sleepTextBox = Create-TextBox -x $actionButtonX -y ($sleepY + 25) -width 150 -height 20
$sleepTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sleepTextBox)
$insertSleepButton = Create-Button -x $actionButtonX -y ($sleepY + 55) -width 100 -height 30 -text "Insert Sleep" -clickAction {
    Write-Host "[EVENT] Insert Sleep Button Clicked" # DEBUG Event Trigger
    try {
        $dur = [int]$sleepTextBox.Text; if($dur -lt 0){throw "Duration must be non-negative."}; # Validate duration
        $act = @{Type="Sleep";Duration=$dur}; # Create sleep action object
        $idx = $listBox.SelectedIndex; # Get selected index
        if($idx -ge 0){ # If item selected, insert after it
            Write-Host "[EVENT]   Inserting Sleep after index $idx" # DEBUG
            $global:Actions.Insert($idx+1,$act); $sel=$idx+1
        } else { # Otherwise, append to the end
            Write-Host "[EVENT]   Appending Sleep to end of list" # DEBUG
            $global:Actions.Add($act); $sel=$global:Actions.Count-1
        }
        Update-ListBox; $listBox.SelectedIndex=$sel # Update display and select new item
    } catch {
        Write-Host "[EVENT]   Error inserting sleep: $($_.Exception.Message)" # DEBUG
        [System.Windows.Forms.MessageBox]::Show("Invalid duration. Please enter a non-negative integer.","Error",0,[System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}
$insertSleepButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($insertSleepButton)

# --- Global Delay Slider --- (Below ListBox, Left Side)
$delayLabel = Create-Label -x 10 -y 210 -width 400 -height 20 -text "Delay between actions: $global:DelayBetweenActions ms"
$delayLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($delayLabel)
$delayTrackBar = Create-TrackBar -x 10 -y 235 -width 400 -min 0 -max 5000 -tick 250 -val $global:DelayBetweenActions -changed {
    # Update global variable and label text when slider value changes
    $global:DelayBetweenActions = $delayTrackBar.Value
    $delayLabel.Text = "Delay between actions: $global:DelayBetweenActions ms"
}
$delayTrackBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($delayTrackBar)


# --- Input Attachment / Edit Action Area --- (Below Delay Slider, Left Side)
$editActionGroupBox = New-Object System.Windows.Forms.GroupBox; $editActionGroupBox.Location = New-Object System.Drawing.Point(10, 270); $editActionGroupBox.Size = New-Object System.Drawing.Size(400, 240); $editActionGroupBox.Text = "Edit Selected Action / Attach Input"; $editActionGroupBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($editActionGroupBox)
# (Controls inside GroupBox - created using factory functions)
$editActionTypeLabel=Create-Label 10 20 80 20 "Type:";$editActionGroupBox.Controls.Add($editActionTypeLabel)
$editActionTypeTextBox=Create-TextBox 90 20 100 20;$editActionTypeTextBox.ReadOnly=$true;$editActionGroupBox.Controls.Add($editActionTypeTextBox)
$editXRelLabel=Create-Label 10 50 80 20 "XRel:";$editActionGroupBox.Controls.Add($editXRelLabel)
$editXRelTextBox=Create-TextBox 90 50 100 20;$editActionGroupBox.Controls.Add($editXRelTextBox)
$editYRelLabel=Create-Label 200 50 80 20 "YRel:";$editActionGroupBox.Controls.Add($editYRelLabel)
$editYRelTextBox=Create-TextBox 280 50 100 20;$editActionGroupBox.Controls.Add($editYRelTextBox)
$editDurationLabel=Create-Label 10 80 80 20 "Duration:";$editDurationLabel.Visible=$false;$editActionGroupBox.Controls.Add($editDurationLabel)
$editDurationTextBox=Create-TextBox 90 80 100 20;$editDurationTextBox.Visible=$false;$editActionGroupBox.Controls.Add($editDurationTextBox)
$editInputLabel=Create-Label 10 110 80 20 "Input:";$editActionGroupBox.Controls.Add($editInputLabel)
$editInputTextBox=Create-TextBox 90 110 290 30;$editInputTextBox.Multiline=$true;$editInputTextBox.ScrollBars="Vertical";$editActionGroupBox.Controls.Add($editInputTextBox)
$editInputInstrLabel=Create-Label 10 145 380 30 "Use '%%VARIABLE_INPUT_N%%' for grid column N.";$editActionGroupBox.Controls.Add($editInputInstrLabel)
$updateActionButton=Create-Button 10 190 100 30 "Update Action" -clickAction {
    Write-Host "[EVENT] Update Action Button Clicked" # DEBUG Event Trigger
    $idx=$listBox.SelectedIndex;
    if($idx -ge 0){
        $a=$global:Actions[$idx];
        Write-Host "[EVENT]   Updating action at index $idx (Type: $($a.Type))" # DEBUG
        # Update properties based on type, with validation
        if($a.Type -eq "Click"){
            try{ $a.XRel=[double]$editXRelTextBox.Text; $a.YRel=[double]$editYRelTextBox.Text; $a.Input=$editInputTextBox.Text; Write-Host "[EVENT]   Click properties updated."} # DEBUG
            catch{[System.Windows.Forms.MessageBox]::Show("Invalid Click props (XRel/YRel must be numbers).","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error); Write-Host "[EVENT]   Error updating Click properties: $($_.Exception.Message)"; return} # DEBUG
        }elseif($a.Type -eq "Sleep"){
            try{$dur=[int]$editDurationTextBox.Text; if($dur -lt 0){throw "Duration must be non-negative."}; $a.Duration=$dur; Write-Host "[EVENT]   Sleep duration updated."} # DEBUG
            catch{[System.Windows.Forms.MessageBox]::Show("Invalid Sleep duration (must be non-negative integer).","Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error); Write-Host "[EVENT]   Error updating Sleep duration: $($_.Exception.Message)"; return} # DEBUG
        }
        Update-ListBox; # Refresh listbox display
        $listBox.SelectedIndex=$idx # Keep item selected
    }else{
        Write-Host "[EVENT]   No item selected to update." # DEBUG
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list first.")
    }
}
$editActionGroupBox.Controls.Add($updateActionButton)
$editCancelButton=Create-Button 120 190 100 30 "Cancel Edit" -clickAction { Write-Host "[EVENT] Cancel Edit Button Clicked"; $listBox.SelectedIndex = -1 } # Deselecting triggers visibility change via ListBox event handler
$editActionGroupBox.Controls.Add($editCancelButton);
# Set initial state for the edit groupbox (disabled until 'Edit Action' is clicked)
$editActionGroupBox.Enabled = $false; $updateActionButton.Enabled = $false; $editCancelButton.Enabled = $false


# --- Save/Load Action Sequence Buttons --- (Right Side, below Sleep)
$saveLoadY = $insertSleepButton.Location.Y + $insertSleepButton.Height + 20 # Position below sleep button
$saveButton = Create-Button -x $actionButtonX -y $saveLoadY -width 150 -height 30 -text "Save Sequence" -clickAction {
    Write-Host "[EVENT] Save Sequence Button Clicked" # DEBUG Event Trigger
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "JSON Sequence (*.json)|*.json|All Files (*.*)|*.*"; $sfd.Title = "Save Action Sequence"; $sfd.DefaultExt = "json"
    if ($sfd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Saving sequence to $($sfd.FileName)" # DEBUG
        try {
            # Convert the actions list to JSON and save to selected file
            ($global:Actions | ConvertTo-Json -Depth 5) | Out-File -FilePath $sfd.FileName -Encoding UTF8
            Write-Host "[EVENT]   Sequence saved successfully." # DEBUG
            [System.Windows.Forms.MessageBox]::Show("Sequence saved successfully to $($sfd.FileName).", "Saved", 0, 'Information')
        } catch {
            Write-Host "[EVENT]   Error saving sequence: $($_.Exception.Message)" # DEBUG
            [System.Windows.Forms.MessageBox]::Show("Error saving sequence: $($_.Exception.Message)", "Error", 0, 'Error')
        }
    } else {
        Write-Host "[EVENT]   Save sequence cancelled by user." # DEBUG
    }
}
$saveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($saveButton)

$loadButton = Create-Button -x $actionButtonX -y ($saveLoadY + 40) -width 150 -height 30 -text "Load Sequence" -clickAction {
    Write-Host "[EVENT] Load Sequence Button Clicked" # DEBUG Event Trigger
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "JSON Sequence (*.json)|*.json|All Files (*.*)|*.*"; $ofd.Title = "Load Action Sequence"
    if ($ofd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Loading sequence from $($ofd.FileName)" # DEBUG
        try {
            # Read the JSON file content
            $json = Get-Content -Path $ofd.FileName -Raw
            # Convert JSON text into PowerShell objects
            Write-Host "[EVENT]   Parsing JSON..." # DEBUG
            $loadedData = $json | ConvertFrom-Json
            # Reset global actions list and populate with loaded data
            Write-Host "[EVENT]   Populating actions list..." # DEBUG
            $global:Actions = [System.Collections.ArrayList]::new()
            if ($loadedData -is [array]) { $global:Actions.AddRange($loadedData) } elseif ($loadedData) { $global:Actions.Add($loadedData) }
            Write-Host "[EVENT]   Actions list populated with $($global:Actions.Count) items." # DEBUG

            # --- Auto-adjust grid columns based on loaded actions ---
            Write-Host "[EVENT]   Scanning loaded actions for max input number..." # DEBUG
            $maxInputNum = 0 # Track the highest %%VARIABLE_INPUT_N%% found
            foreach ($action in $global:Actions) {
                # Check if action has an 'Input' property and it contains a placeholder pattern
                if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input -match $global:VariableInputScanPattern) {
                    # Find all placeholder matches within the input string
                    $matches = $action.Input | Select-String -Pattern $global:VariableInputScanPattern -AllMatches
                    if ($matches) {
                        foreach ($match in $matches.Matches) {
                            try {
                                # Extract the number N from the placeholder
                                $num = [int]$match.Groups[1].Value
                                # Update the maximum number found so far
                                if ($num -gt $maxInputNum) { $maxInputNum = $num }
                            } catch { Write-Warning "Could not parse input number from '$($match.Groups[1].Value)' in action input '$($action.Input)'" }
                        }
                    }
                }
            }
            Write-Host "[EVENT]   Max Input Number found in loaded sequence: $maxInputNum" # DEBUG
            $currentColCount = $dataGridView.Columns.Count
            # If the sequence requires more columns than the grid currently has
            if ($maxInputNum -gt $currentColCount) {
                Write-Host "[EVENT]   Adding columns to grid (up to Input $maxInputNum)..." # DEBUG
                $dataGridView.SuspendLayout() # Suspend grid layout updates for performance
                $columnsAdded = $false
                # Loop from the next needed column up to the maximum required
                for ($i = $currentColCount + 1; $i -le $maxInputNum; $i++) {
                    # Use the helper function to add the column
                    if (Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd $i) {
                        $columnsAdded = $true # Flag that at least one column was added
                    }
                }
                $dataGridView.ResumeLayout() # Resume grid layout updates
                # Notify the user if columns were added
                if ($columnsAdded) {
                    [System.Windows.Forms.MessageBox]::Show("Grid columns automatically adjusted to match loaded sequence.", "Grid Updated", 0, 'Information')
                }
            } else {
                 Write-Host "[EVENT]   Grid columns already sufficient ($currentColCount >= $maxInputNum)." # DEBUG
            }
            # --- End auto-adjust grid columns ---

            Update-ListBox # Refresh the action list display
            Write-Host "[EVENT]   Sequence loaded and processed successfully." # DEBUG
            [System.Windows.Forms.MessageBox]::Show("Sequence loaded successfully from $($ofd.FileName).", "Loaded", 0, 'Information')

        } catch {
            # Provide more detailed error info in the console
            Write-Host "[EVENT]   Error loading sequence: $($_.Exception.Message)" -ForegroundColor Red # DEBUG
            Write-Host "[EVENT]   StackTrace: $($_.ScriptStackTrace)" -ForegroundColor Yellow # DEBUG More detail
            [System.Windows.Forms.MessageBox]::Show("Error loading or processing sequence: $($_.Exception.Message)", "Load Error", 0, 'Error')
            # Clear actions list if loading failed critically to avoid inconsistent state
            $global:Actions = [System.Collections.ArrayList]::new()
            Update-ListBox
        }
    } else {
        Write-Host "[EVENT]   Load sequence cancelled by user." # DEBUG
    }
}
$loadButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($loadButton)


# --- Repeat Count --- (Right Side)
$repeatY = $loadButton.Location.Y + $loadButton.Height + 10
$repeatLabel = Create-Label -x $actionButtonX -y $repeatY -width 150 -height 20 -text "Repeat Entire Process:"
$repeatLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($repeatLabel)
$repeatTextBox = Create-TextBox -x ($actionButtonX + 155) -y $repeatY -width 55 -height 20 # Adjusted X/Width for alignment
$repeatTextBox.Text = "1"; $repeatTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($repeatTextBox)


# --- Run Automation Button --- (Right Side)
$runButtonY = $repeatY + $repeatTextBox.Height + 10
$runButton = Create-Button -x $actionButtonX -y $runButtonY -width 150 -height 30 -text "Run Automation"
# --- Use the CORRECTED Run Button logic that handles combined input ---
$runButton.Add_Click({
    Write-Host "[EVENT] Run Automation Button Clicked" # DEBUG Event Trigger
    # Determine the repeat count
    $processRepeatCount = 1; if ($repeatTextBox.Text.Trim() -ne "") { try { $processRepeatCount = [int]$repeatTextBox.Text; if($processRepeatCount -le 0) { throw } } catch { [void][System.Windows.Forms.MessageBox]::Show("Invalid Repeat (>0). Using 1.","Warn",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning); $processRepeatCount = 1; $repeatTextBox.Text = "1" } }
    # Check grid data rows (excluding the potential 'new row')
    $dataRowCount = $dataGridView.Rows.Count; if ($dataGridView.AllowUserToAddRows) { $dataRowCount-- }

    if ($dataRowCount -gt 0) { # Grid Mode: Run using data from the grid
        Write-Host "[RUN]   Running in Grid Mode ($dataRowCount rows, $processRepeatCount repeats)" # DEBUG
        if ([System.Windows.Forms.MessageBox]::Show("Run sequence using $dataRowCount data row(s) from grid? Process repeats $processRepeatCount time(s).", "Confirm Run", 'OKCancel', 'Question') -ne 'OK') { Write-Host "[RUN]   Run cancelled by user."; return }
        Write-Host "[RUN] Starting automation with grid data..."
        # Outer loop for repeating the entire process
        for ($pr = 1; $pr -le $processRepeatCount; $pr++) {
            Write-Host "[RUN]  Starting Process Repeat #$pr"
            $rowIndex = 0
            # Middle loop iterating through each data row in the grid
            foreach ($gridRow in $dataGridView.Rows) {
                if ($gridRow.IsNewRow) { continue } # Skip the blank 'new row' at the end
                $rowIndex++
                Write-Host "[RUN]   Running sequence for Grid Row #$rowIndex"
                $actionIndex = 0
                # Inner loop iterating through the action sequence
                foreach ($action in $global:Actions) {
                    $actionIndex++
                    Write-Host "[RUN]    Executing Action #$actionIndex : $($action.Type)"
                    # --- Click Action Logic ---
                    if ($action.Type -eq "Click") {
                        Move-And-Click -xRel $action.XRel -yRel $action.YRel # Perform move and click
                        Start-Sleep -Milliseconds $global:DelayBeforeClick # Wait before potential input
                        # Check if this action has input defined
                        if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                            $processedInput = $action.Input # Start with the raw input string
                            # Check if the input contains any variable placeholder pattern
                            if ($processedInput -match $global:VariableInputScanPattern) {
                                # Loop through grid columns to find matching placeholders
                                for ($n = 1; $n -le $dataGridView.Columns.Count; $n++) {
                                    $placeholder = "%%VARIABLE_INPUT_$n%%" # Construct placeholder string
                                    $colName = "Input$n" # Construct corresponding column name
                                    # Check if placeholder exists in string AND column exists in grid
                                    if (($processedInput -like "*$placeholder*") -and $dataGridView.Columns.Contains($colName)) {
                                        $valueToInsert = [string]$gridRow.Cells[$colName].Value # Get value from grid cell
                                        # Replace all occurrences of the placeholder with the grid value
                                        $processedInput = $processedInput -replace [regex]::Escape($placeholder), $valueToInsert
                                        Write-Host "[RUN]     Substituted '$placeholder' with '$valueToInsert'"
                                    }
                                }
                            }
                            # Send the final processed string (after potential substitutions) if not empty
                            if ($processedInput.Trim() -ne "") {
                                Write-Host "[RUN]     Sending processed input: '$processedInput'"
                                [System.Windows.Forms.SendKeys]::SendWait($processedInput) # Send keys
                                Start-Sleep -Milliseconds $global:DelayAfterClick # Wait after sending keys
                            }
                        }
                        Start-Sleep -Milliseconds $global:DelayBetweenActions # Wait after the whole Click action
                    # --- Sleep Action Logic ---
                    } elseif ($action.Type -eq "Sleep") {
                        Write-Host "[RUN]     Sleeping for $($action.Duration) ms"
                        Start-Sleep -Milliseconds $action.Duration # Perform sleep
                        Start-Sleep -Milliseconds $global:DelayBetweenActions # Wait after the Sleep action
                    }
                } # End actions loop
                 Write-Host "[RUN]   Sequence complete for Grid Row #$rowIndex"
            } # End grid rows loop
             Write-Host "[RUN]  Process Repeat #$pr Complete"
        } # End process repeats loop
        Write-Host "[RUN] Automation complete (Grid Mode)."
        [void][System.Windows.Forms.MessageBox]::Show("Automation complete.", "Finished", 0, 'Information')

    } else { # Normal Mode: Run without grid data
        Write-Host "[RUN]   Running in Normal Mode (No grid data, $processRepeatCount repeats)" # DEBUG
        if ([System.Windows.Forms.MessageBox]::Show("No data in grid. Run sequence normally? Process repeats $processRepeatCount time(s).", "Confirm Run", 'OKCancel', 'Question') -ne 'OK') { Write-Host "[RUN]   Run cancelled by user."; return }
        Write-Host "[RUN] Starting automation normally (no grid data)..."
        # Outer loop for repeating the sequence
        for ($r = 1; $r -le $processRepeatCount; $r++) {
            Write-Host "[RUN]  Starting Repeat #$r"
            $actionIndex = 0
            # Inner loop for actions
            foreach ($action in $global:Actions) {
                $actionIndex++
                Write-Host "[RUN]   Executing Action #$actionIndex : $($action.Type)"
                # --- Click Action Logic (Normal Mode) ---
                if ($action.Type -eq "Click") {
                    Move-And-Click -xRel $action.XRel -yRel $action.YRel
                    Start-Sleep -Milliseconds $global:DelayBeforeClick
                    # Check for input, but only send if it's NOT a placeholder
                    if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                        if ($action.Input -match $global:VariableInputScanPattern) {
                            # Warn if a placeholder is found but we're not using grid data
                            Write-Warning "[RUN]    Action #$actionIndex marked for variable input ('$($action.Input)'), but no grid data is present. Input step skipped."
                        } else {
                            # Send static input
                            Write-Host "[RUN]     Sending static input: $($action.Input)"
                            [System.Windows.Forms.SendKeys]::SendWait($action.Input)
                            Start-Sleep -Milliseconds $global:DelayAfterClick
                        }
                    }
                    Start-Sleep -Milliseconds $global:DelayBetweenActions # Wait after click action
                # --- Sleep Action Logic (Normal Mode) ---
                } elseif ($action.Type -eq "Sleep") {
                    Write-Host "[RUN]     Sleeping for $($action.Duration) ms"
                    Start-Sleep -Milliseconds $action.Duration
                    Start-Sleep -Milliseconds $global:DelayBetweenActions # Wait after sleep action
                }
            } # End actions loop
             Write-Host "[RUN]  Repeat #$r Complete"
        } # End repeats loop
         Write-Host "[RUN] Automation complete (Normal Mode)."
         [void][System.Windows.Forms.MessageBox]::Show("Automation complete.", "Finished", 0, 'Information')
    }
})
$runButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($runButton)


# --- DataGridView Section --- (Below Edit GroupBox)
$gridLabelY = $editActionGroupBox.Location.Y + $editActionGroupBox.Height + 10
$dataGridViewLabel = Create-Label -x 10 -y $gridLabelY -width 400 -height 20 -text "Variable Input Data Grid (Each row is one run):"
$dataGridViewLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($dataGridViewLabel)

# --- Grid Control Buttons --- (Below Grid Label)
$gridButtonY = $dataGridViewLabel.Location.Y + $dataGridViewLabel.Height + 5
$gridButtonWidth = 140 # Width for grid buttons
$addColButton = Create-Button -x 10 -y $gridButtonY -width $gridButtonWidth -height 30 -text "Add Input Column" -clickAction {
    Write-Host "[EVENT] Add Input Column Button Clicked" # DEBUG Event Trigger
    # Use helper function to add the next sequential column
    if (Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd ($dataGridView.Columns.Count + 1)) {
        [System.Windows.Forms.MessageBox]::Show("Column 'Input$($dataGridView.Columns.Count)' added.", "Column Added", 0, 'Information')
    } # No message if column already existed (shouldn't happen with this logic)
}
$addColButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($addColButton)

$importCsvButton = Create-Button -x (10 + $gridButtonWidth + 5) -y $gridButtonY -width $gridButtonWidth -height 30 -text "Import Data (CSV)" -clickAction {
    Write-Host "[EVENT] Import Data (CSV) Button Clicked" # DEBUG Event Trigger
    $confirm = [System.Windows.Forms.MessageBox]::Show("Clear existing grid data before importing?", "Confirm Import", 'YesNoCancel', 'Question')
    if ($confirm -eq 'Cancel') { Write-Host "[EVENT]   Import cancelled by user."; return }
    if ($confirm -eq 'Yes') { Write-Host "[EVENT]   Clearing existing grid data."; $dataGridView.Rows.Clear() } # Clear grid rows if user confirmed
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"; $ofd.Title = "Select CSV (Must have InputN headers)"
    if ($ofd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Importing from $($ofd.FileName)" # DEBUG
        try {
            # Use Semicolon delimiter based on user context (Austria)
            $imported = Import-Csv -Path $ofd.FileName -Delimiter ';'
            if ($imported) { # Check if import returned data
                Write-Host "[EVENT]   CSV read successfully, $($imported.Count) rows found. Populating grid..." # DEBUG
                $dataGridView.SuspendLayout() # Suspend layout for performance
                # Iterate through each row imported from the CSV
                foreach ($row in $imported) {
                    $idx = $dataGridView.Rows.Add() # Add a new row to the grid
                    $gridRow = $dataGridView.Rows[$idx] # Get the newly added grid row object
                    # Iterate through the columns currently in the grid
                    foreach ($col in $dataGridView.Columns) {
                        # Check if the imported data row has a property matching the grid column name
                        if ($row.PSObject.Properties.Match($col.Name).Count -gt 0) {
                            # Assign the value from the CSV to the grid cell (as string)
                            $gridRow.Cells[$col.Name].Value = [string]$row.$($col.Name)
                        } else {
                            # Optional: Log if CSV header missing for existing grid column
                            # Write-Warning "CSV row missing expected header '$($col.Name)'"
                            $gridRow.Cells[$col.Name].Value = [string]::Empty # Set to empty if missing in CSV
                        }
                    }
                }
                $dataGridView.ResumeLayout() # Resume layout updates
                Write-Host "[EVENT]   Grid populated." # DEBUG
                [System.Windows.Forms.MessageBox]::Show("Data imported successfully from CSV.", "Import Complete", 0, 'Information')
            } else {
                Write-Host "[EVENT]   CSV file empty or unreadable." # DEBUG
                [System.Windows.Forms.MessageBox]::Show("CSV file appears empty or could not be read.", "Import Warning", 0, 'Warning')
            }
        } catch {
            Write-Host "[EVENT]   Error importing CSV: $($_.Exception.Message)" # DEBUG
            [System.Windows.Forms.MessageBox]::Show("Error importing data from CSV: $($_.Exception.Message)", "Import Error", 0, 'Error')
        }
    } else {
        Write-Host "[EVENT]   Import file selection cancelled." # DEBUG
    }
}
$importCsvButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($importCsvButton)

$saveGridButton = Create-Button -x (10 + ($gridButtonWidth + 5) * 2) -y $gridButtonY -width $gridButtonWidth -height 30 -text "Save Grid Data (CSV)" -clickAction {
    Write-Host "[EVENT] Save Grid Data Button Clicked" # DEBUG Event Trigger
     $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"; $sfd.Title = "Save Grid Data As CSV"; $sfd.DefaultExt = "csv"; $sfd.FileName = "grid_data.csv"
     if ($sfd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Saving grid data to $($sfd.FileName)" # DEBUG
        try {
            $dataToExport = [System.Collections.Generic.List[PSCustomObject]]::new() # List to hold data rows
            $columnNames = $dataGridView.Columns | ForEach-Object { $_.Name } # Get current grid column names
            if ($columnNames.Count -eq 0) { throw "Grid has no columns to export." } # Cannot export if no columns

            # Iterate through grid rows (excluding the 'new row')
            Write-Host "[EVENT]   Extracting data from grid..." # DEBUG
            foreach ($gridRow in $dataGridView.Rows) {
                if ($gridRow.IsNewRow) { continue }
                $rowObject = [ordered]@{} # Use ordered dictionary for consistent column order in CSV
                $isEmptyRow = $true # Flag to track if row contains any data
                # Iterate through the defined column names
                foreach ($colName in $columnNames) {
                    $cellValue = [string]$gridRow.Cells[$colName].Value # Get cell value as string
                    $rowObject[$colName] = $cellValue # Add to ordered dictionary
                    # Check if the cell value is non-empty
                    if (-not [string]::IsNullOrEmpty($cellValue)) { $isEmptyRow = $false }
                }
                # Only add the row to the export list if it wasn't completely empty
                if (-not $isEmptyRow) { $dataToExport.Add([PSCustomObject]$rowObject) }
            }
            Write-Host "[EVENT]   Extracted $($dataToExport.Count) non-empty rows." # DEBUG

            # Check if there is any data to export
            if ($dataToExport.Count -gt 0) {
                # Export the collected data using semicolon delimiter and UTF8 encoding
                Write-Host "[EVENT]   Exporting data to CSV..." # DEBUG
                $dataToExport | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Host "[EVENT]   Grid data saved successfully." # DEBUG
                [System.Windows.Forms.MessageBox]::Show("Grid data saved successfully to $($sfd.FileName).", "Success", 0, 'Information')
            } else {
                Write-Host "[EVENT]   Grid contains no data to save." # DEBUG
                [System.Windows.Forms.MessageBox]::Show("Grid contains no data to save (excluding empty rows).", "Empty Grid", 0, 'Information')
            }
        } catch {
            Write-Host "[EVENT]   Error saving grid data: $($_.Exception.Message)" # DEBUG
            [System.Windows.Forms.MessageBox]::Show("Error saving grid data: $($_.Exception.Message)", "Error", 0, 'Error')
        }
     } else {
         Write-Host "[EVENT]   Save grid data cancelled by user." # DEBUG
     }
}
$saveGridButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($saveGridButton)


# --- DataGridView Control --- (Below Grid Buttons)
$dataGridViewY = $gridButtonY + $addColButton.Height + 5 # Y position below the buttons
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, $dataGridViewY)
# Calculate initial size dynamically based on form client area
# Use explicit integer casting and minimum size checks
$dataGridViewHeight = [int]$form.ClientSize.Height - [int]$dataGridViewY - 20 # Leave 20px padding at bottom
$dataGridViewWidth = [int]$form.ClientSize.Width - 20 # Leave 10px padding on each side
if ($dataGridViewWidth -lt 100) { $dataGridViewWidth = 100 } # Enforce minimum width
if ($dataGridViewHeight -lt 100) { $dataGridViewHeight = 100 } # Enforce minimum height
$dataGridView.Size = New-Object System.Drawing.Size($dataGridViewWidth, $dataGridViewHeight)
# Anchor the grid to all sides so it resizes with the form
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
# Configure grid properties
$dataGridView.AllowUserToAddRows = $true       # Allow adding new rows via the last row
$dataGridView.AllowUserToDeleteRows = $true   # Allow deleting rows (select row header and press Delete key)
$dataGridView.AutoSizeColumnsMode = 'Fill'     # Make columns fill the available width
$dataGridView.ColumnHeadersHeightSizeMode = 'AutoSize' # Adjust header height automatically

# Add initial columns using the helper function
Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd 1 | Out-Null
Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd 2 | Out-Null
# Add the configured grid to the form's controls
$form.Controls.Add($dataGridView)


# --- KeyDown Handler (F9 for Capture / Escape for Cancel Edit) ---
$form.Add_KeyDown({
    param($sender, $e) # Event arguments
    # F9 Key: Capture mouse coordinates
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F9) {
        $pos = [System.Windows.Forms.Cursor]::Position; $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        if ($screen.Width -le 0 -or $screen.Height -le 0) { return } # Safety check
        $xr = $pos.X / $screen.Width; $yr = $pos.Y / $screen.Height # Calculate relative coords
        # Create new action object (Hashtable)
        $act = @{Type="Click"; XRel=[Math]::Round($xr,4); YRel=[Math]::Round($yr,4); Input=""}
        $global:Actions.Add($act) # Add to global list
        Update-ListBox # Refresh display
        $listBox.SelectedIndex = $global:Actions.Count - 1 # Select the newly added action
        $e.Handled = $true; $e.SuppressKeyPress = $true; # Mark event as handled, suppress system beep
    }
    # Escape Key: Cancel editing state
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        # Only act if the edit groupbox is currently enabled
        if ($editActionGroupBox.Enabled) {
            Write-Host "[EVENT] Escape key pressed - cancelling edit." # DEBUG
            $listBox.SelectedIndex = -1 # Deselecting triggers the ListBox event handler to disable edit box
            $e.Handled = $true; $e.SuppressKeyPress = $true; # Mark event as handled
        }
    }
})

# --- ListBox Selection Changed Handler ---
# Handles enabling/disabling the edit groupbox when selection changes
$listBox.Add_SelectedIndexChanged({
    if ($listBox.SelectedIndex -lt 0) { # If no item is selected
        # Disable the edit groupbox and its buttons
        if ($editActionGroupBox.Enabled) { # Only write host if state changes
            Write-Host "[DEBUG] ListBox selection cleared - disabling Edit GroupBox." # DEBUG
            $editActionGroupBox.Enabled = $false
            $updateActionButton.Enabled = $false
            $editCancelButton.Enabled = $false
        }
    }
    # Note: We don't automatically enable the box on selection; user must click 'Edit Action' button
})

# --- Show the Form ---
# Display the form window and wait for it to close
Write-Host "[DEBUG] Showing main form..." # DEBUG
[void]$form.ShowDialog()
#endregion Form Creation and Controls

# Final message when script exits
Write-Host "Script finished."

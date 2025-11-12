#region Assembly Loading
# Load necessary .NET assemblies for GUI elements
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#endregion Assembly Loading

#region Win32 API Definitions
# P/Invoke signatures for window management and mouse control
$win32Signature = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Win32API {
    // Get the handle of the foreground (active) window
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    // Get the dimensions and position of a window
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    // Get the process ID that owns a window
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    
    // Get the window title text
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    
    // Check if a window is visible
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    // Find a window by class name and/or window name
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    // Enumerate all top-level windows
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
}

// Structure to hold window rectangle coordinates
[StructLayout(LayoutKind.Sequential)]
public struct RECT {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
}
"@

# Check if Win32API type already exists
$existingWin32Type = [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {
    $_.GetType("Win32API", $false)
} | Where-Object { $_ -ne $null }
if (-not $existingWin32Type) {
    Add-Type -TypeDefinition $win32Signature -PassThru | Out-Null
}
#endregion Win32 API Definitions

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
$existingType = [AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object {
    $_.GetType("MouseSimulator", $false)
} | Where-Object { $_ -ne $null }
if (-not $existingType) {
    Add-Type -TypeDefinition $mouseSignature -PassThru | Out-Null
}

# Function to simulate a left mouse button click (down then up)
function Click-MouseButton {
    [MouseSimulator]::mouse_event([MouseSimulator]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    [MouseSimulator]::mouse_event([MouseSimulator]::MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
}

# Function to move the mouse cursor to absolute screen coordinates and then click
function Move-And-Click {
    param(
        [Parameter(Mandatory=$true)]
        [int]$x,  # Absolute X screen coordinate
        [Parameter(Mandatory=$true)]
        [int]$y   # Absolute Y screen coordinate
    )
    # Set the cursor position
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
    # Short delay to allow UI to potentially update before clicking
    Start-Sleep -Milliseconds 100
    # Perform the click
    Click-MouseButton
}

# Helper function to get window info from a window handle
function Get-WindowInfo {
    param([IntPtr]$hWnd)
    
    if ($hWnd -eq [IntPtr]::Zero) { return $null }
    
    # Get window title
    $titleBuilder = New-Object System.Text.StringBuilder 256
    [Win32API]::GetWindowText($hWnd, $titleBuilder, 256) | Out-Null
    $title = $titleBuilder.ToString()
    
    # Get process ID and name
    $processId = 0
    [Win32API]::GetWindowThreadProcessId($hWnd, [ref]$processId) | Out-Null
    $processName = ""
    try {
        $proc = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($proc) { $processName = $proc.ProcessName }
    } catch { }
    
    # Get window rectangle
    $rect = New-Object RECT
    [Win32API]::GetWindowRect($hWnd, [ref]$rect) | Out-Null
    
    return [PSCustomObject]@{
        Handle = [string]$hWnd
        Title = $title
        ProcessName = $processName
        ProcessId = $processId
        Left = $rect.Left
        Top = $rect.Top
        Width = ($rect.Right - $rect.Left)
        Height = ($rect.Bottom - $rect.Top)
    }
}

# Function to find a window by process name and/or title pattern
function Find-WindowByInfo {
    param(
        [string]$ProcessName,
        [string]$TitlePattern
    )
    
    $foundWindow = $null
    $callback = {
        param([IntPtr]$hWnd, [IntPtr]$lParam)
        
        if ([Win32API]::IsWindowVisible($hWnd)) {
            $info = Get-WindowInfo -hWnd $hWnd
            if ($info) {
                $processMatch = (-not $ProcessName) -or ($info.ProcessName -eq $ProcessName)
                $titleMatch = (-not $TitlePattern) -or ($info.Title -like "*$TitlePattern*")
                
                if ($processMatch -and $titleMatch -and $info.Width -gt 0 -and $info.Height -gt 0) {
                    $script:foundWindow = $info
                    return $false  # Stop enumeration
                }
            }
        }
        return $true  # Continue enumeration
    }
    
    $delegateType = [Win32API].GetNestedType('EnumWindowsProc', [System.Reflection.BindingFlags]::Public)
    $delegate = [System.Delegate]::CreateDelegate($delegateType, $callback.GetType().GetMethod('Invoke', [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::Public))
    [Win32API]::EnumWindows($delegate, [IntPtr]::Zero) | Out-Null
    
    return $script:foundWindow
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
    $listBox.BeginUpdate()
    $listBox.Items.Clear()
    for ($i = 0; $i -lt $global:Actions.Count; $i++) {
        $action = $global:Actions[$i]
        $displayText = "Action $([int]($i+1)): $($action.Type)"
        if ($action.Type -eq "Click") {
            # Show window-relative info if available
            if ($action.PSObject.Properties.Name -contains 'WindowProcess' -and $action.WindowProcess) {
                $displayText += " [$($action.WindowProcess)]"
            }
            $displayText += " at Win($($action.WinX), $($action.WinY))"
            # Add associated input text if it exists and is not empty
            if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                 $displayText += " - Input: $($action.Input)"
            }
        } elseif ($action.Type -eq "Sleep") {
            $displayText += " for $($action.Duration) ms"
        }
        $listBox.Items.Add($displayText) | Out-Null
    }
    $listBox.EndUpdate()
    Write-Host "[DEBUG] ListBox Updated with $($global:Actions.Count) items."
}

# Helper function to add a new InputN column to the DataGridView if it doesn't already exist
function Add-GridInputColumn {
    param(
        [Parameter(Mandatory=$true)]
        [ref]$dataGridViewRef,
        [Parameter(Mandatory=$true)]
        [int]$columnNumberToAdd
    )
    $dataGridView = $dataGridViewRef.Value
    $colName = "Input$columnNumberToAdd"
    $headerText = "Input $columnNumberToAdd"
    if (-not $dataGridView.Columns.Contains($colName)) {
        $newCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $newCol.Name = $colName
        $newCol.HeaderText = $headerText
        $dataGridView.Columns.Add($newCol) | Out-Null
        Write-Host "[DEBUG] Helper Function: Added grid column: $headerText"
        return $true
    }
    return $false
}

# Factory function for creating Label controls
function Create-Label { param([int]$x,[int]$y,[int]$width,[int]$height,[string]$text) $l=New-Object System.Windows.Forms.Label; $l.Location=New-Object System.Drawing.Point($x,$y); $l.Size=New-Object System.Drawing.Size($width,$height); $l.Text=$text; return $l }

# Factory function for creating Button controls
function Create-Button {
    param(
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [string]$text,
        [scriptblock]$clickAction
    )
    Write-Host "[DEBUG] Creating Button: '$text'"
    $b = New-Object System.Windows.Forms.Button
    $b.Location = New-Object System.Drawing.Point($x,$y)
    $b.Size = New-Object System.Drawing.Size($width,$height)
    $b.Text = $text
    if ($clickAction -ne $null) {
        Write-Host "[DEBUG]   Attempting to attach Click Action for '$text'"
        try {
            $b.Add_Click($clickAction)
            Write-Host "[DEBUG]   Click Action attached successfully for '$text'."
        } catch {
             Write-Host "[DEBUG]   ERROR attaching Click Action for '$text': $($_.Exception.Message)"
        }
    } else {
         Write-Host "[DEBUG]   WARNING: No Click Action provided for '$text'."
    }
    return $b
}

# Factory function for creating TextBox controls
function Create-TextBox { param([int]$x,[int]$y,[int]$width,[int]$height) $tb=New-Object System.Windows.Forms.TextBox; $tb.Location=New-Object System.Drawing.Point($x,$y); $tb.Size=New-Object System.Drawing.Size($width,$height); return $tb }
# Factory function for creating TrackBar controls
function Create-TrackBar { param([int]$x,[int]$y,[int]$width,[int]$min,[int]$max,[int]$tick,[int]$val,[scriptblock]$changed) $tk=New-Object System.Windows.Forms.TrackBar;$tk.Location=New-Object System.Drawing.Point($x,$y);$tk.Width=$width;$tk.Minimum=$min;$tk.Maximum=$max;$tk.TickFrequency=$tick;$tk.Value=$val;$tk.Add_ValueChanged($changed);return $tk }

#endregion GUI Functions and Elements


#region Form Creation and Controls
# Create the main application window (Form)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Steiner's Automatisierer QoL (Window-Relative v3)"
$form.Size = New-Object System.Drawing.Size(750, 900)
$form.StartPosition = "CenterScreen"
$form.KeyPreview = $true

# --- Top Instruction Label ---
$instructionLabel = Create-Label -x 10 -y 10 -width ($form.ClientSize.Width - 20) -height 30 -text "F9: Capture (window-relative). Esc: Cancel Edit. Use %%VARIABLE_INPUT_N%% for grid data."
$instructionLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($instructionLabel)

# --- Action Sequence ListBox --- (Left Side)
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,50)
$listBox.Size = New-Object System.Drawing.Size(400,150)
$listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listBox)

# --- Action List Controls --- (Right Side)
$actionButtonX = 420
$actionButtonWidth = 100
$actionButtonGap = 5

# Column 1 of action buttons (Copy, Clear All)
$copyButton = Create-Button -x $actionButtonX -y 50 -width $actionButtonWidth -height 30 -text "Copy Action" -clickAction {
    Write-Host "[EVENT] Copy Action Button Clicked"
    $selectedIndex = $listBox.SelectedIndex
    if ($selectedIndex -ge 0) {
        Write-Host "[EVENT]   Item selected at index: $selectedIndex"
        $originalAction = $global:Actions[$selectedIndex]
        Write-Host "[EVENT]   Original Action: $($originalAction | ConvertTo-Json -Depth 2 -Compress)"
        $newAction = ($originalAction | ConvertTo-Json -Depth 5 | ConvertFrom-Json)
        $global:Actions.Insert($selectedIndex + 1, $newAction)
        Write-Host "[EVENT]   Action copied and inserted at index $($selectedIndex + 1)"
        Update-ListBox
        $listBox.SelectedIndex = $selectedIndex + 1
    } else {
        Write-Host "[EVENT]   No item selected to copy"
        [System.Windows.Forms.MessageBox]::Show("Please select an action in the list to copy.", "No Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
}
$copyButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($copyButton)

$clearButton = Create-Button -x $actionButtonX -y 90 -width $actionButtonWidth -height 30 -text "Clear All" -clickAction {
    Write-Host "[EVENT] Clear All Button Clicked"
    if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear all actions in the sequence?", "Confirm Clear", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning) -eq 'Yes') {
        Write-Host "[EVENT]   Clearing all actions..."
        $global:Actions.Clear()
        Update-ListBox
        Write-Host "[EVENT]   Actions cleared."
    } else {
        Write-Host "[EVENT]   Clear cancelled by user."
    }
}
$clearButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($clearButton)

# Column 2 of action buttons (Move Up, Move Down, Delete, Edit)
$col2X = $actionButtonX + $actionButtonWidth + $actionButtonGap
$moveUpButton = Create-Button -x $col2X -y 50 -width $actionButtonWidth -height 30 -text "Move Up" -clickAction {
    Write-Host "[EVENT] Move Up Button Clicked"
    $index = $listBox.SelectedIndex
    if ($index -gt 0) {
        Write-Host "[EVENT]   Moving item at index $index up"
        $itemToMove = $global:Actions[$index]
        $global:Actions.RemoveAt($index)
        $global:Actions.Insert($index - 1, $itemToMove)
        Update-ListBox
        $listBox.SelectedIndex = $index - 1
    } else {
        Write-Host "[EVENT]   Cannot move item up (not selected or already at top)"
    }
}
$moveUpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($moveUpButton)

$moveDownButton = Create-Button -x $col2X -y 90 -width $actionButtonWidth -height 30 -text "Move Down" -clickAction {
    Write-Host "[EVENT] Move Down Button Clicked"
    $index = $listBox.SelectedIndex
    if ($index -ge 0 -and $index -lt ($global:Actions.Count - 1)) {
        Write-Host "[EVENT]   Moving item at index $index down"
        $itemToMove = $global:Actions[$index]
        $global:Actions.RemoveAt($index)
        $global:Actions.Insert($index + 1, $itemToMove)
        Update-ListBox
        $listBox.SelectedIndex = $index + 1
    } else {
         Write-Host "[EVENT]   Cannot move item down (not selected or already at bottom)"
    }
}
$moveDownButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($moveDownButton)

$deleteButton = Create-Button -x $col2X -y 130 -width $actionButtonWidth -height 30 -text "Delete" -clickAction {
    Write-Host "[EVENT] Delete Button Clicked"
    $index = $listBox.SelectedIndex
    if ($index -ge 0) {
        Write-Host "[EVENT]   Deleting item at index $index"
        $global:Actions.RemoveAt($index)
        Update-ListBox
    } else {
        Write-Host "[EVENT]   No item selected to delete"
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list to delete.")
    }
}
$deleteButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($deleteButton)

$editActionButton = Create-Button -x $col2X -y 170 -width $actionButtonWidth -height 30 -text "Edit Action" -clickAction {
    Write-Host "[EVENT] Edit Action Button Clicked"
    $idx = $listBox.SelectedIndex;
    if ($idx -ge 0) {
        Write-Host "[EVENT]   Editing item at index $idx"
        $selectedAction = $global:Actions[$idx];
        $editActionGroupBox.Enabled = $true
        $updateActionButton.Enabled = $true
        $editCancelButton.Enabled = $true
        $editActionTypeTextBox.Text = $selectedAction.Type;
        Write-Host "[EVENT]   Action Type: $($selectedAction.Type)"
        if ($selectedAction.Type -eq "Click") {
             Write-Host "[EVENT]   Populating Click fields..."
             $editXRelLabel.Visible = $true
             $editXRelTextBox.Visible = $true
             $editYRelLabel.Visible = $true
             $editYRelTextBox.Visible = $true
             $editInputLabel.Visible = $true
             $editInputTextBox.Visible = $true
             $editDurationLabel.Visible = $false
             $editDurationTextBox.Visible = $false;
             # Populate click controls - use WinX/WinY (window-relative)
             $editXRelTextBox.Text = $selectedAction.WinX
             $editYRelTextBox.Text = $selectedAction.WinY
             # Safely get Input property (might not exist on old saves)
             $inputValue = ""
             if ($selectedAction.PSObject.Properties.Name -contains 'Input') {
                 $inputValue = $selectedAction.Input
             }
             $editInputTextBox.Text = $inputValue
        } elseif ($selectedAction.Type -eq "Sleep") {
             Write-Host "[EVENT]   Populating Sleep fields..."
             $editDurationLabel.Visible = $true
             $editDurationTextBox.Visible = $true;
             $editXRelLabel.Visible = $false
             $editXRelTextBox.Visible = $false
             $editYRelLabel.Visible = $false
             $editYRelTextBox.Visible = $false
             $editInputLabel.Visible = $false
             $editInputTextBox.Visible = $false
             $editDurationTextBox.Text = $selectedAction.Duration
        }
        Write-Host "[EVENT]   Edit GroupBox enabled and populated."
    } else {
        Write-Host "[EVENT]   No item selected to edit."
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list to edit.")
    }
}
$editActionButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($editActionButton)

# --- Insert Sleep --- (Right Side, below action buttons)
$sleepY = 210
$sleepLabel = Create-Label -x $actionButtonX -y $sleepY -width 150 -height 20 -text "Sleep Duration (ms):"
$sleepLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sleepLabel)
$sleepTextBox = Create-TextBox -x $actionButtonX -y ($sleepY + 25) -width 150 -height 20
$sleepTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($sleepTextBox)
$insertSleepButton = Create-Button -x $actionButtonX -y ($sleepY + 55) -width 100 -height 30 -text "Insert Sleep" -clickAction {
    Write-Host "[EVENT] Insert Sleep Button Clicked"
    try {
        $dur = [int]$sleepTextBox.Text; if($dur -lt 0){throw "Duration must be non-negative."};
        $act = [PSCustomObject]@{Type="Sleep";Duration=$dur};
        $idx = $listBox.SelectedIndex;
        if($idx -ge 0){
            Write-Host "[EVENT]   Inserting Sleep after index $idx"
            $global:Actions.Insert($idx+1,$act); $sel=$idx+1
        } else {
            Write-Host "[EVENT]   Appending Sleep to end of list"
            $global:Actions.Add($act) | Out-Null; $sel=$global:Actions.Count-1
        }
        Update-ListBox; $listBox.SelectedIndex=$sel
    } catch {
        Write-Host "[EVENT]   Error inserting sleep: $($_.Exception.Message)"
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
    $global:DelayBetweenActions = $delayTrackBar.Value
    $delayLabel.Text = "Delay between actions: $global:DelayBetweenActions ms"
}
$delayTrackBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($delayTrackBar)


# --- Input Attachment / Edit Action Area --- (Below Delay Slider, Left Side)
$editActionGroupBox = New-Object System.Windows.Forms.GroupBox; $editActionGroupBox.Location = New-Object System.Drawing.Point(10, 270); $editActionGroupBox.Size = New-Object System.Drawing.Size(400, 240); $editActionGroupBox.Text = "Edit Selected Action / Attach Input"; $editActionGroupBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left; $form.Controls.Add($editActionGroupBox)
# (Controls inside GroupBox)
$editActionTypeLabel=Create-Label 10 20 80 20 "Type:";$editActionGroupBox.Controls.Add($editActionTypeLabel)
$editActionTypeTextBox=Create-TextBox 90 20 100 20;$editActionTypeTextBox.ReadOnly=$true;$editActionGroupBox.Controls.Add($editActionTypeTextBox)
$editXRelLabel=Create-Label 10 50 80 20 "WinX:";$editActionGroupBox.Controls.Add($editXRelLabel)
$editXRelTextBox=Create-TextBox 90 50 100 20;$editActionGroupBox.Controls.Add($editXRelTextBox)
$editYRelLabel=Create-Label 200 50 80 20 "WinY:";$editActionGroupBox.Controls.Add($editYRelLabel)
$editYRelTextBox=Create-TextBox 280 50 100 20;$editActionGroupBox.Controls.Add($editYRelTextBox)
$editDurationLabel=Create-Label 10 80 80 20 "Duration:";$editDurationLabel.Visible=$false;$editActionGroupBox.Controls.Add($editDurationLabel)
$editDurationTextBox=Create-TextBox 90 80 100 20;$editDurationTextBox.Visible=$false;$editActionGroupBox.Controls.Add($editDurationTextBox)
$editInputLabel=Create-Label 10 110 80 20 "Input:";$editActionGroupBox.Controls.Add($editInputLabel)
$editInputTextBox=Create-TextBox 90 110 290 30;$editInputTextBox.Multiline=$true;$editInputTextBox.ScrollBars="Vertical";$editActionGroupBox.Controls.Add($editInputTextBox)
$editInputInstrLabel=Create-Label 10 145 380 30 "Use '%%VARIABLE_INPUT_N%%' for grid column N.";$editActionGroupBox.Controls.Add($editInputInstrLabel)
$updateActionButton=Create-Button 10 190 100 30 "Update Action" -clickAction {
    Write-Host "[EVENT] Update Action Button Clicked"
    $idx=$listBox.SelectedIndex;
    if($idx -ge 0){
        $a=$global:Actions[$idx];
        Write-Host "[EVENT]   Updating action at index $idx (Type: $($a.Type))"
        if($a.Type -eq "Click"){
            try{ 
                $a.WinX=[int]$editXRelTextBox.Text
                $a.WinY=[int]$editYRelTextBox.Text
                $a.Input=$editInputTextBox.Text
                Write-Host "[EVENT]   Click properties updated."
            }
            catch{
                [System.Windows.Forms.MessageBox]::Show("Invalid Click props (WinX/WinY must be integers).","Error",0,16)
                Write-Host "[EVENT]   Error updating Click properties: $($_.Exception.Message)"
                return
            }
        }elseif($a.Type -eq "Sleep"){
            try{
                $dur=[int]$editDurationTextBox.Text
                if($dur -lt 0){throw "Duration must be non-negative."}
                $a.Duration=$dur
                Write-Host "[EVENT]   Sleep duration updated."
            }
            catch{
                [System.Windows.Forms.MessageBox]::Show("Invalid Sleep duration (must be non-negative integer).","Error",0,16)
                Write-Host "[EVENT]   Error updating Sleep duration: $($_.Exception.Message)"
                return
            }
        }
        Update-ListBox;
        $listBox.SelectedIndex=$idx
    }else{
        Write-Host "[EVENT]   No item selected to update."
        [System.Windows.Forms.MessageBox]::Show("Select an action in the list first.")
    }
}
$editActionGroupBox.Controls.Add($updateActionButton)
$editCancelButton=Create-Button 120 190 100 30 "Cancel Edit" -clickAction { Write-Host "[EVENT] Cancel Edit Button Clicked"; $listBox.SelectedIndex = -1 }
$editActionGroupBox.Controls.Add($editCancelButton);
$editActionGroupBox.Enabled = $false; $updateActionButton.Enabled = $false; $editCancelButton.Enabled = $false


# --- Save/Load Action Sequence Buttons --- (Right Side, below Sleep)
$saveLoadY = $insertSleepButton.Location.Y + $insertSleepButton.Height + 20
$saveButton = Create-Button -x $actionButtonX -y $saveLoadY -width 150 -height 30 -text "Save Sequence" -clickAction {
    Write-Host "[EVENT] Save Sequence Button Clicked"
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "JSON Sequence (*.json)|*.json|All Files (*.*)|*.*"; $sfd.Title = "Save Action Sequence"; $sfd.DefaultExt = "json"
    if ($sfd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Saving sequence to $($sfd.FileName)"
        try {
            ($global:Actions | ConvertTo-Json -Depth 5) | Out-File -FilePath $sfd.FileName -Encoding UTF8
            Write-Host "[EVENT]   Sequence saved successfully."
            [System.Windows.Forms.MessageBox]::Show("Sequence saved successfully to $($sfd.FileName).", "Saved", 0, 'Information')
        } catch {
            Write-Host "[EVENT]   Error saving sequence: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error saving sequence: $($_.Exception.Message)", "Error", 0, 'Error')
        }
    } else {
        Write-Host "[EVENT]   Save sequence cancelled by user."
    }
}
$saveButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($saveButton)

$loadButton = Create-Button -x $actionButtonX -y ($saveLoadY + 40) -width 150 -height 30 -text "Load Sequence" -clickAction {
    Write-Host "[EVENT] Load Sequence Button Clicked"
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "JSON Sequence (*.json)|*.json|All Files (*.*)|*.*"; $ofd.Title = "Load Action Sequence"
    if ($ofd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Loading sequence from $($ofd.FileName)"
        try {
            $json = Get-Content -Path $ofd.FileName -Raw
            Write-Host "[EVENT]   Parsing JSON..."
            $loadedData = $json | ConvertFrom-Json
            Write-Host "[EVENT]   Populating actions list..."
            $global:Actions = [System.Collections.ArrayList]::new()
            if ($loadedData -is [array]) { $global:Actions.AddRange($loadedData) } elseif ($loadedData) { $global:Actions.Add($loadedData) | Out-Null }
            Write-Host "[EVENT]   Actions list populated with $($global:Actions.Count) items."

            # --- Auto-adjust grid columns based on loaded actions ---
            Write-Host "[EVENT]   Scanning loaded actions for max input number..."
            $maxInputNum = 0
            foreach ($action in $global:Actions) {
                if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input -match $global:VariableInputScanPattern) {
                    $matches = $action.Input | Select-String -Pattern $global:VariableInputScanPattern -AllMatches
                    if ($matches) {
                        foreach ($match in $matches.Matches) {
                            try {
                                $num = [int]$match.Groups[1].Value
                                if ($num -gt $maxInputNum) { $maxInputNum = $num }
                            } catch { Write-Warning "Could not parse input number from '$($match.Groups[1].Value)' in action input '$($action.Input)'" }
                        }
                    }
                }
            }
            Write-Host "[EVENT]   Max Input Number found in loaded sequence: $maxInputNum"
            $currentColCount = $dataGridView.Columns.Count
            if ($maxInputNum -gt $currentColCount) {
                Write-Host "[EVENT]   Adding columns to grid (up to Input $maxInputNum)..."
                $dataGridView.SuspendLayout()
                $columnsAdded = $false
                for ($i = $currentColCount + 1; $i -le $maxInputNum; $i++) {
                    if (Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd $i) {
                        $columnsAdded = $true
                    }
                }
                $dataGridView.ResumeLayout()
                if ($columnsAdded) {
                    [System.Windows.Forms.MessageBox]::Show("Grid columns automatically adjusted to match loaded sequence.", "Grid Updated", 0, 'Information')
                }
            } else {
                 Write-Host "[EVENT]   Grid columns already sufficient ($currentColCount >= $maxInputNum)."
            }

            Update-ListBox
            Write-Host "[EVENT]   Sequence loaded and processed successfully."
            [System.Windows.Forms.MessageBox]::Show("Sequence loaded successfully from $($ofd.FileName).", "Loaded", 0, 'Information')

        } catch {
            Write-Host "[EVENT]   Error loading sequence: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "[EVENT]   StackTrace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("Error loading or processing sequence: $($_.Exception.Message)", "Load Error", 0, 'Error')
            $global:Actions = [System.Collections.ArrayList]::new()
            Update-ListBox
        }
    } else {
        Write-Host "[EVENT]   Load sequence cancelled by user."
    }
}
$loadButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($loadButton)


# --- Repeat Count --- (Right Side)
$repeatY = $loadButton.Location.Y + $loadButton.Height + 10
$repeatLabel = Create-Label -x $actionButtonX -y $repeatY -width 150 -height 20 -text "Repeat Entire Process:"
$repeatLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($repeatLabel)
$repeatTextBox = Create-TextBox -x ($actionButtonX + 155) -y $repeatY -width 55 -height 20
$repeatTextBox.Text = "1"; $repeatTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($repeatTextBox)


# --- Run Automation Button --- (Right Side)
$runButtonY = $repeatY + $repeatTextBox.Height + 10
$runButton = Create-Button -x $actionButtonX -y $runButtonY -width 150 -height 30 -text "Run Automation"
$runButton.Add_Click({
    Write-Host "[EVENT] Run Automation Button Clicked"
    $processRepeatCount = 1
    if ($repeatTextBox.Text.Trim() -ne "") {
        try {
            $processRepeatCount = [int]$repeatTextBox.Text
            if($processRepeatCount -le 0) { throw }
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show("Invalid Repeat (>0). Using 1.","Warn",0,48)
            $processRepeatCount = 1
            $repeatTextBox.Text = "1"
        }
    }
    $dataRowCount = $dataGridView.Rows.Count
    if ($dataGridView.AllowUserToAddRows) { $dataRowCount-- }

    if ($dataRowCount -gt 0) {
        Write-Host "[RUN]   Running in Grid Mode ($dataRowCount rows, $processRepeatCount repeats)"
        if ([System.Windows.Forms.MessageBox]::Show("Run sequence using $dataRowCount data row(s) from grid? Process repeats $processRepeatCount time(s).", "Confirm Run", 'OKCancel', 'Question') -ne 'OK') {
            Write-Host "[RUN]   Run cancelled by user."
            return
        }
        Write-Host "[RUN] Starting automation with grid data..."
        for ($pr = 1; $pr -le $processRepeatCount; $pr++) {
            Write-Host "[RUN]  Starting Process Repeat #$pr"
            $rowIndex = 0
            foreach ($gridRow in $dataGridView.Rows) {
                if ($gridRow.IsNewRow) { continue }
                $rowIndex++
                Write-Host "[RUN]   Running sequence for Grid Row #$rowIndex"
                $actionIndex = 0
                foreach ($action in $global:Actions) {
                    $actionIndex++
                    Write-Host "[RUN]    Executing Action #$actionIndex : $($action.Type)"
                    if ($action.Type -eq "Click") {
                        # Calculate screen position from window-relative coordinates
                        $targetX = 0
                        $targetY = 0
                        
                        # Try to find the target window if window info was saved
                        $windowFound = $false
                        if ($action.PSObject.Properties.Name -contains 'WindowProcess' -and $action.WindowProcess) {
                            $windowInfo = Find-WindowByInfo -ProcessName $action.WindowProcess -TitlePattern $action.WindowTitle
                            if ($windowInfo -and $windowInfo.Width -gt 0 -and $windowInfo.Height -gt 0) {
                                $targetX = $windowInfo.Left + $action.WinX
                                $targetY = $windowInfo.Top + $action.WinY
                                $windowFound = $true
                                Write-Host "[RUN]     Found window [$($action.WindowProcess)], clicking at screen ($targetX, $targetY)"
                            }
                        }
                        
                        # Fallback: use screen-relative if window not found or no window info stored
                        if (-not $windowFound) {
                            $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
                            $targetX = $action.WinX
                            $targetY = $action.WinY
                            Write-Host "[RUN]     Window info not available or not found, using absolute coordinates ($targetX, $targetY)"
                        }
                        
                        Move-And-Click -x $targetX -y $targetY
                        Start-Sleep -Milliseconds $global:DelayBeforeClick
                        
                        if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                            $processedInput = $action.Input
                            if ($processedInput -match $global:VariableInputScanPattern) {
                                for ($n = 1; $n -le $dataGridView.Columns.Count; $n++) {
                                    $placeholder = "%%VARIABLE_INPUT_$n%%"
                                    $colName = "Input$n"
                                    if (($processedInput -like "*$placeholder*") -and $dataGridView.Columns.Contains($colName)) {
                                        $valueToInsert = [string]$gridRow.Cells[$colName].Value
                                        $processedInput = $processedInput -replace [regex]::Escape($placeholder), $valueToInsert
                                        Write-Host "[RUN]     Substituted '$placeholder' with '$valueToInsert'"
                                    }
                                }
                            }
                            if ($processedInput.Trim() -ne "") {
                                Write-Host "[RUN]     Sending processed input: '$processedInput'"
                                [System.Windows.Forms.SendKeys]::SendWait($processedInput)
                                Start-Sleep -Milliseconds $global:DelayAfterClick
                            }
                        }
                        Start-Sleep -Milliseconds $global:DelayBetweenActions
                    } elseif ($action.Type -eq "Sleep") {
                        Write-Host "[RUN]     Sleeping for $($action.Duration) ms"
                        Start-Sleep -Milliseconds $action.Duration
                        Start-Sleep -Milliseconds $global:DelayBetweenActions
                    }
                }
                 Write-Host "[RUN]   Sequence complete for Grid Row #$rowIndex"
            }
             Write-Host "[RUN]  Process Repeat #$pr Complete"
        }
        Write-Host "[RUN] Automation complete (Grid Mode)."
        [void][System.Windows.Forms.MessageBox]::Show("Automation complete.", "Finished", 0, 'Information')

    } else {
        Write-Host "[RUN]   Running in Normal Mode (No grid data, $processRepeatCount repeats)"
        if ([System.Windows.Forms.MessageBox]::Show("No data in grid. Run sequence normally? Process repeats $processRepeatCount time(s).", "Confirm Run", 'OKCancel', 'Question') -ne 'OK') {
            Write-Host "[RUN]   Run cancelled by user."
            return
        }
        Write-Host "[RUN] Starting automation normally (no grid data)..."
        for ($r = 1; $r -le $processRepeatCount; $r++) {
            Write-Host "[RUN]  Starting Repeat #$r"
            $actionIndex = 0
            foreach ($action in $global:Actions) {
                $actionIndex++
                Write-Host "[RUN]   Executing Action #$actionIndex : $($action.Type)"
                if ($action.Type -eq "Click") {
                    $targetX = 0
                    $targetY = 0
                    
                    $windowFound = $false
                    if ($action.PSObject.Properties.Name -contains 'WindowProcess' -and $action.WindowProcess) {
                        $windowInfo = Find-WindowByInfo -ProcessName $action.WindowProcess -TitlePattern $action.WindowTitle
                        if ($windowInfo -and $windowInfo.Width -gt 0 -and $windowInfo.Height -gt 0) {
                            $targetX = $windowInfo.Left + $action.WinX
                            $targetY = $windowInfo.Top + $action.WinY
                            $windowFound = $true
                            Write-Host "[RUN]     Found window [$($action.WindowProcess)], clicking at screen ($targetX, $targetY)"
                        }
                    }
                    
                    if (-not $windowFound) {
                        $targetX = $action.WinX
                        $targetY = $action.WinY
                        Write-Host "[RUN]     Window info not available or not found, using absolute coordinates ($targetX, $targetY)"
                    }
                    
                    Move-And-Click -x $targetX -y $targetY
                    Start-Sleep -Milliseconds $global:DelayBeforeClick
                    
                    if ($action.PSObject.Properties.Name -contains 'Input' -and $action.Input -and $action.Input.Trim() -ne "") {
                        if ($action.Input -match $global:VariableInputScanPattern) {
                            Write-Warning "[RUN]    Action #$actionIndex marked for variable input ('$($action.Input)'), but no grid data is present. Input step skipped."
                        } else {
                            Write-Host "[RUN]     Sending static input: $($action.Input)"
                            [System.Windows.Forms.SendKeys]::SendWait($action.Input)
                            Start-Sleep -Milliseconds $global:DelayAfterClick
                        }
                    }
                    Start-Sleep -Milliseconds $global:DelayBetweenActions
                } elseif ($action.Type -eq "Sleep") {
                    Write-Host "[RUN]     Sleeping for $($action.Duration) ms"
                    Start-Sleep -Milliseconds $action.Duration
                    Start-Sleep -Milliseconds $global:DelayBetweenActions
                }
            }
             Write-Host "[RUN]  Repeat #$r Complete"
        }
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
$gridButtonWidth = 140
$addColButton = Create-Button -x 10 -y $gridButtonY -width $gridButtonWidth -height 30 -text "Add Input Column" -clickAction {
    Write-Host "[EVENT] Add Input Column Button Clicked"
    if (Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd ($dataGridView.Columns.Count + 1)) {
        [System.Windows.Forms.MessageBox]::Show("Column 'Input$($dataGridView.Columns.Count)' added.", "Column Added", 0, 'Information')
    }
}
$addColButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($addColButton)

$importCsvButton = Create-Button -x (10 + $gridButtonWidth + 5) -y $gridButtonY -width $gridButtonWidth -height 30 -text "Import Data (CSV)" -clickAction {
    Write-Host "[EVENT] Import Data (CSV) Button Clicked"
    $confirm = [System.Windows.Forms.MessageBox]::Show("Clear existing grid data before importing?", "Confirm Import", 'YesNoCancel', 'Question')
    if ($confirm -eq 'Cancel') { Write-Host "[EVENT]   Import cancelled by user."; return }
    if ($confirm -eq 'Yes') { Write-Host "[EVENT]   Clearing existing grid data."; $dataGridView.Rows.Clear() }
    $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"; $ofd.Title = "Select CSV (Must have InputN headers)"
    if ($ofd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Importing from $($ofd.FileName)"
        try {
            $imported = Import-Csv -Path $ofd.FileName -Delimiter ';'
            if ($imported) {
                Write-Host "[EVENT]   CSV read successfully, $($imported.Count) rows found. Populating grid..."
                $dataGridView.SuspendLayout()
                foreach ($row in $imported) {
                    $idx = $dataGridView.Rows.Add()
                    $gridRow = $dataGridView.Rows[$idx]
                    foreach ($col in $dataGridView.Columns) {
                        if ($row.PSObject.Properties.Match($col.Name).Count -gt 0) {
                            $gridRow.Cells[$col.Name].Value = [string]$row.$($col.Name)
                        } else {
                            $gridRow.Cells[$col.Name].Value = [string]::Empty
                        }
                    }
                }
                $dataGridView.ResumeLayout()
                Write-Host "[EVENT]   Grid populated."
                [System.Windows.Forms.MessageBox]::Show("Data imported successfully from CSV.", "Import Complete", 0, 'Information')
            } else {
                Write-Host "[EVENT]   CSV file empty or unreadable."
                [System.Windows.Forms.MessageBox]::Show("CSV file appears empty or could not be read.", "Import Warning", 0, 'Warning')
            }
        } catch {
            Write-Host "[EVENT]   Error importing CSV: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error importing data from CSV: $($_.Exception.Message)", "Import Error", 0, 'Error')
        }
    } else {
        Write-Host "[EVENT]   Import file selection cancelled."
    }
}
$importCsvButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($importCsvButton)

$saveGridButton = Create-Button -x (10 + ($gridButtonWidth + 5) * 2) -y $gridButtonY -width $gridButtonWidth -height 30 -text "Save Grid Data (CSV)" -clickAction {
    Write-Host "[EVENT] Save Grid Data Button Clicked"
     $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"; $sfd.Title = "Save Grid Data As CSV"; $sfd.DefaultExt = "csv"; $sfd.FileName = "grid_data.csv"
     if ($sfd.ShowDialog() -eq 'OK') {
        Write-Host "[EVENT]   Saving grid data to $($sfd.FileName)"
        try {
            $dataToExport = [System.Collections.Generic.List[PSCustomObject]]::new()
            $columnNames = $dataGridView.Columns | ForEach-Object { $_.Name }
            if ($columnNames.Count -eq 0) { throw "Grid has no columns to export." }

            Write-Host "[EVENT]   Extracting data from grid..."
            foreach ($gridRow in $dataGridView.Rows) {
                if ($gridRow.IsNewRow) { continue }
                $rowObject = [ordered]@{}
                $isEmptyRow = $true
                foreach ($colName in $columnNames) {
                    $cellValue = [string]$gridRow.Cells[$colName].Value
                    $rowObject[$colName] = $cellValue
                    if (-not [string]::IsNullOrEmpty($cellValue)) { $isEmptyRow = $false }
                }
                if (-not $isEmptyRow) { $dataToExport.Add([PSCustomObject]$rowObject) }
            }
            Write-Host "[EVENT]   Extracted $($dataToExport.Count) non-empty rows."

            if ($dataToExport.Count -gt 0) {
                Write-Host "[EVENT]   Exporting data to CSV..."
                $dataToExport | Export-Csv -Path $sfd.FileName -NoTypeInformation -Delimiter ';' -Encoding UTF8
                Write-Host "[EVENT]   Grid data saved successfully."
                [System.Windows.Forms.MessageBox]::Show("Grid data saved successfully to $($sfd.FileName).", "Success", 0, 'Information')
            } else {
                Write-Host "[EVENT]   Grid contains no data to save."
                [System.Windows.Forms.MessageBox]::Show("Grid contains no data to save (excluding empty rows).", "Empty Grid", 0, 'Information')
            }
        } catch {
            Write-Host "[EVENT]   Error saving grid data: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Error saving grid data: $($_.Exception.Message)", "Error", 0, 'Error')
        }
     } else {
         Write-Host "[EVENT]   Save grid data cancelled by user."
     }
}
$saveGridButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($saveGridButton)


# --- DataGridView Control --- (Below Grid Buttons)
$dataGridViewY = $gridButtonY + $addColButton.Height + 5
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, $dataGridViewY)
$dataGridViewHeight = [int]$form.ClientSize.Height - [int]$dataGridViewY - 20
$dataGridViewWidth = [int]$form.ClientSize.Width - 20
if ($dataGridViewWidth -lt 100) { $dataGridViewWidth = 100 }
if ($dataGridViewHeight -lt 100) { $dataGridViewHeight = 100 }
$dataGridView.Size = New-Object System.Drawing.Size($dataGridViewWidth, $dataGridViewHeight)
$dataGridView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGridView.AllowUserToAddRows = $true
$dataGridView.AllowUserToDeleteRows = $true
$dataGridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.ColumnHeadersHeightSizeMode = 'AutoSize'

Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd 1 | Out-Null
Add-GridInputColumn -dataGridViewRef ([ref]$dataGridView) -columnNumberToAdd 2 | Out-Null
$form.Controls.Add($dataGridView)


# --- KeyDown Handler (F9 for Capture / Escape for Cancel Edit) ---
$form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'F9') {
        # Get cursor position and foreground window info
        $pos = [System.Windows.Forms.Cursor]::Position
        $fgWindow = [Win32API]::GetForegroundWindow()
        
        # Get window information
        $windowInfo = Get-WindowInfo -hWnd $fgWindow
        
        if ($windowInfo -and $windowInfo.Width -gt 0 -and $windowInfo.Height -gt 0) {
            # Calculate window-relative coordinates
            $winX = $pos.X - $windowInfo.Left
            $winY = $pos.Y - $windowInfo.Top
            
            # Clamp to window bounds (safety check)
            if ($winX -lt 0) { $winX = 0 }
            if ($winY -lt 0) { $winY = 0 }
            if ($winX -gt $windowInfo.Width) { $winX = $windowInfo.Width }
            if ($winY -gt $windowInfo.Height) { $winY = $windowInfo.Height }
            
            Write-Host "[F9] Captured click: Screen($($pos.X),$($pos.Y)) -> Window-relative($winX,$winY) in [$($windowInfo.ProcessName)] '$($windowInfo.Title)'"
            
            # Create action with window-relative coordinates and window info
            $act = [PSCustomObject]@{
                Type = "Click"
                WinX = $winX
                WinY = $winY
                WindowProcess = $windowInfo.ProcessName
                WindowTitle = $windowInfo.Title
                Input = ""
            }
        } else {
            # Fallback: no window found or invalid window, use absolute screen coordinates
            Write-Host "[F9] Captured click: No valid window detected, using absolute screen coordinates ($($pos.X),$($pos.Y))"
            $act = [PSCustomObject]@{
                Type = "Click"
                WinX = $pos.X
                WinY = $pos.Y
                WindowProcess = ""
                WindowTitle = ""
                Input = ""
            }
        }
        
        $global:Actions.Add($act) | Out-Null
        Update-ListBox
        $listBox.SelectedIndex = $global:Actions.Count - 1
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    }
    if ($e.KeyCode -eq 'Escape') {
        if ($editActionGroupBox.Enabled) {
            Write-Host "[EVENT] Escape key pressed - cancelling edit."
            $listBox.SelectedIndex = -1
            $e.Handled = $true
            $e.SuppressKeyPress = $true
        }
    }
})

# --- ListBox Selection Changed Handler ---
$listBox.Add_SelectedIndexChanged({
    if ($listBox.SelectedIndex -lt 0) {
        if ($editActionGroupBox.Enabled) {
            Write-Host "[DEBUG] ListBox selection cleared - disabling Edit GroupBox."
            $editActionGroupBox.Enabled = $false
            $updateActionButton.Enabled = $false
            $editCancelButton.Enabled = $false
        }
    }
})

# --- Show the Form ---
Write-Host "[DEBUG] Showing main form..."
[void]$form.ShowDialog()
#endregion Form Creation and Controls

Write-Host "Script finished."

# DYNAMIC_FINAL_FINAL_refactored_pwsh5.ps1
# -----------------------------------------------------------------------------
# Windows PowerShell 5.1–compatible refactor
#  - Removed PowerShell 7 operators (??, ?., ?:)
#  - Replaced Math.Clamp with custom Clamp()
#  - Fixed Edit flow & list refresh
#  - Window‑relative coordinates persist & survive window moves/resizes
# -----------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Native interop -----------------------------------------------------------
$pinvoke = @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class Native {
  [DllImport("user32.dll", SetLastError=true)]
  public static extern IntPtr GetForegroundWindow();

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

  [DllImport("user32.dll", CharSet=CharSet.Auto)]
  public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

  [DllImport("user32.dll", SetLastError=true)]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

  [StructLayout(LayoutKind.Sequential)]
  public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
  [StructLayout(LayoutKind.Sequential)]
  public struct POINT { public int X; public int Y; }
}

public static class MouseSim {
  [DllImport("user32.dll", CallingConvention=CallingConvention.StdCall)]
  public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
  public const uint MOUSEEVENTF_MOVE      = 0x0001;
  public const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
  public const uint MOUSEEVENTF_LEFTUP    = 0x0004;
}
"@
Add-Type -TypeDefinition $pinvoke -PassThru | Out-Null

# --- Globals -----------------------------------------------------------------
$global:Actions = New-Object System.Collections.ArrayList
$global:DelayBeforeClick    = 150
$global:DelayAfterClick     = 200
$global:DelayBetweenActions = 350
$global:SequencePath = $null

# Helper Clamp for .NET Framework
function Clamp([double]$v,[double]$min,[double]$max){ if($v -lt $min){return $min} if($v -gt $max){return $max} return $v }

# --- Helpers -----------------------------------------------------------------
function Get-ActiveWindowInfo {
  $hwnd = [Native]::GetForegroundWindow()
  if($hwnd -eq [IntPtr]::Zero){ return $null }
  $titleSb = New-Object System.Text.StringBuilder 512
  [void][Native]::GetWindowText($hwnd, $titleSb, $titleSb.Capacity)
  $pid = 0
  [void][Native]::GetWindowThreadProcessId($hwnd, [ref]$pid)
  $pname = $null
  try { $proc = Get-Process -Id $pid -ErrorAction Stop; $pname = $proc.ProcessName } catch {}
  [pscustomobject]@{ Hwnd=$hwnd; Title=$titleSb.ToString(); Process=$pname }
}

function Get-WindowClientRectScreen([IntPtr]$hwnd){
  if($hwnd -eq [IntPtr]::Zero){ return $null }
  $rc = New-Object Native+RECT
  if(-not [Native]::GetClientRect($hwnd, [ref]$rc)){ return $null }
  # client origin to screen
  $pt = New-Object Native+POINT
  $pt.X = 0; $pt.Y=0
  if(-not [Native]::ClientToScreen($hwnd, [ref]$pt)){ return $null }
  # convert to screen space rect
  return [pscustomobject]@{ X=$pt.X; Y=$pt.Y; W=($rc.Right - $rc.Left); H=($rc.Bottom - $rc.Top) }
}

function Move-Cursor([int]$x,[int]$y){ [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y) }
function Click-Left(){ [MouseSim]::mouse_event([MouseSim]::MOUSEEVENTF_LEFTDOWN,0,0,0,0); [MouseSim]::mouse_event([MouseSim]::MOUSEEVENTF_LEFTUP,0,0,0,0) }

function Resolve-TargetHwnd($action){
  if($action.Mode -ne 'Window'){ return [Native]::GetForegroundWindow() }
  $procName = $null; $titleNeedle = $null
  if($action.Target){ $procName = $action.Target.Process; $titleNeedle = $action.Target.Title }
  if([string]::IsNullOrEmpty($procName)) { return [Native]::GetForegroundWindow() }
  $candidates = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq $procName }
  foreach($p in $candidates){
    try{
      $h = $p.MainWindowHandle
      if($h -eq 0){ continue }
      $sb = New-Object System.Text.StringBuilder 512
      [void][Native]::GetWindowText($h, $sb, $sb.Capacity)
      if([string]::IsNullOrEmpty($titleNeedle) -or ($sb.ToString().ToLower().Contains($titleNeedle.ToLower()))){ return $h }
    } catch {}
  }
  return [Native]::GetForegroundWindow()
}

function To-ScreenPoint($action){
  if($action.Type -ne 'Click'){ return $null }
  if($action.Mode -eq 'Screen'){
    $scr = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $x = [int]([Math]::Round($scr.Width  * $action.XRel))
    $y = [int]([Math]::Round($scr.Height * $action.YRel))
    return @{X=$x;Y=$y}
  } else {
    $hwnd = Resolve-TargetHwnd $action
    $rc = Get-WindowClientRectScreen $hwnd
    if(-not $rc){ return $null }
    $x = [int]([Math]::Round($rc.X + $rc.W * $action.XRel))
    $y = [int]([Math]::Round($rc.Y + $rc.H * $action.YRel))
    return @{X=$x;Y=$y}
  }
}

function New-ClickAction([string]$mode,[double]$xRel,[double]$yRel,[string]$input,[object]$target){
  $h = @{ Type='Click'; Mode=$mode; XRel=[Math]::Round($xRel,4); YRel=[Math]::Round($yRel,4); Input=''; Target=$null }
  if(-not [string]::IsNullOrEmpty($input)){ $h.Input = $input }
  if($target -ne $null){ $h.Target = $target }
  return $h
}

function New-SleepAction([int]$ms){ @{ Type='Sleep'; Duration=([Math]::Max(0,[int]$ms)) } }

function Action-ToString($a){
  if($a.Type -eq 'Click'){
    $scope = 'Screen'
    if($a.Mode -eq 'Window'){
      $p = '?'
      if($a.Target -and $a.Target.Process){ $p = $a.Target.Process }
      $scope = "Win:" + $p
    }
    $inp = ''
    if($a.Input){ $inp = $a.Input.Replace("`n","↵") }
    return ("Click [" + $scope + "] XRel=" + $a.XRel + " YRel=" + $a.YRel + " Input='" + $inp + "'")
  } else { return ("Sleep " + $a.Duration + " ms") }
}

# --- Persistence --------------------------------------------------------------
function Save-Sequence([string]$path){ ($global:Actions | ConvertTo-Json -Depth 6) | Set-Content -Path $path -Encoding UTF8 }
function Load-Sequence([string]$path){
  $global:Actions.Clear() | Out-Null
  $loaded = Get-Content -Path $path -Raw | ConvertFrom-Json
  foreach($a in $loaded){ [void]$global:Actions.Add([hashtable]$a.PSObject.Copy()) }
}

# --- Playback ----------------------------------------------------------------
function Play-Sequence{
  foreach($a in $global:Actions){
    switch($a.Type){
      'Click' {
        $pt = To-ScreenPoint $a
        if(-not $pt){ [System.Windows.Forms.MessageBox]::Show("Could not resolve click point."); break }
        Move-Cursor $pt.X $pt.Y
        Start-Sleep -Milliseconds $global:DelayBeforeClick
        Click-Left
        if($a.Input){ [System.Windows.Forms.SendKeys]::SendWait($a.Input) }
        Start-Sleep -Milliseconds $global:DelayAfterClick
      }
      'Sleep' { Start-Sleep -Milliseconds ([int]$a.Duration) }
    }
    Start-Sleep -Milliseconds $global:DelayBetweenActions
  }
}

# --- UI ----------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'UI Clicker – window‑relative'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(900, 560)
$form.TopMost = $false

$list = New-Object System.Windows.Forms.ListBox
$list.Location = New-Object System.Drawing.Point(10,10)
$list.Size = New-Object System.Drawing.Size(540, 500)
$list.Anchor = 'Top,Left,Bottom'
$form.Controls.Add($list)

function Refresh-List { $list.Items.Clear(); $global:Actions | ForEach-Object { [void]$list.Items.Add((Action-ToString $_)) } }

# Capture area: choose Screen vs Window
$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = 'Capture / Insert'
$grp.Location = New-Object System.Drawing.Point(560,10)
$grp.Size = New-Object System.Drawing.Size(320, 190)
$grp.Anchor = 'Top,Right'
$form.Controls.Add($grp)

$radWin = New-Object System.Windows.Forms.RadioButton; $radWin.Text='Window‑relative'; $radWin.Location='10,20'; $radWin.Checked=$true
$radScr = New-Object System.Windows.Forms.RadioButton; $radScr.Text='Screen‑relative'; $radScr.Location='160,20'
$grp.Controls.AddRange(@($radWin,$radScr))

$lblT = New-Object System.Windows.Forms.Label; $lblT.Text='Target window:'; $lblT.Location='10,50'; $lblT.AutoSize=$true
$grp.Controls.Add($lblT)
$combo = New-Object System.Windows.Forms.ComboBox; $combo.Location='10,70'; $combo.Width=300; $combo.DropDownStyle='DropDownList'; $combo.DisplayMember='Display'
$grp.Controls.Add($combo)

function Refresh-WindowList{
  $items = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 } | ForEach-Object {
    $sb = New-Object System.Text.StringBuilder 512
    [void][Native]::GetWindowText($_.MainWindowHandle, $sb, $sb.Capacity)
    $obj = New-Object PSObject -Property @{
      Display = ("(" + $_.ProcessName + ") — " + $sb.ToString())
      Process = $_.ProcessName
      Title   = $sb.ToString()
    }
    $obj
  }
  $combo.Items.Clear(); [void]$combo.Items.AddRange(($items | ForEach-Object { $_ }))
  if($combo.Items.Count -gt 0){ $combo.SelectedIndex = 0 }
}

$btnRefreshWins = New-Object System.Windows.Forms.Button; $btnRefreshWins.Text='↻'; $btnRefreshWins.Location='280,46'; $btnRefreshWins.Width=30
$btnRefreshWins.Add_Click({ Refresh-WindowList })
$grp.Controls.Add($btnRefreshWins)

$btnAddClick = New-Object System.Windows.Forms.Button; $btnAddClick.Text='Add Click (current mouse)'; $btnAddClick.Location='10,110'; $btnAddClick.Width=300
$btnAddClick.Add_Click({
  $mode = 'Screen'; if($radWin.Checked){ $mode = 'Window' }
  $cursor = [System.Windows.Forms.Cursor]::Position
  if($mode -eq 'Screen'){
    $scr = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $xr = Clamp (($cursor.X / [double]$scr.Width)  ) 0 1
    $yr = Clamp (($cursor.Y / [double]$scr.Height) ) 0 1
    $act = New-ClickAction $mode $xr $yr '' $null
  } else {
    $sel = $combo.SelectedItem
    if(-not $sel){ [System.Windows.Forms.MessageBox]::Show('Select a target window first.'); return }
    $proc = $sel.Process; $title = $sel.Title
    $p = Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq $proc } | Select-Object -First 1
    if(-not $p){ [System.Windows.Forms.MessageBox]::Show('Process not found.'); return }
    $rc = Get-WindowClientRectScreen $p.MainWindowHandle
    if(-not $rc){ [System.Windows.Forms.MessageBox]::Show('Client rect not available.'); return }
    $xr = Clamp ((($cursor.X - $rc.X) / [double]$rc.W)) 0 1
    $yr = Clamp ((($cursor.Y - $rc.Y) / [double]$rc.H)) 0 1
    $target = @{ Title=$title; Process=$proc }
    $act = New-ClickAction $mode $xr $yr '' $target
  }
  [void]$global:Actions.Add($act); Refresh-List;  $list.SelectedIndex = $global:Actions.Count-1
})
$grp.Controls.Add($btnAddClick)

$btnAddSleep = New-Object System.Windows.Forms.Button; $btnAddSleep.Text='Insert Sleep 500 ms'; $btnAddSleep.Location='10,145'; $btnAddSleep.Width=300
$btnAddSleep.Add_Click({
  $sel = $list.SelectedIndex
  $a = New-SleepAction 500
  if($sel -ge 0){ $global:Actions.Insert($sel+1,$a) } else { [void]$global:Actions.Add($a) }
  Refresh-List; $list.SelectedIndex = ([Math]::Min($global:Actions.Count-1, [Math]::Max(0,$sel+1)))
})
$grp.Controls.Add($btnAddSleep)

# Edit panel
$edit = New-Object System.Windows.Forms.GroupBox; $edit.Text='Edit selected'; $edit.Location='560,210'; $edit.Size='320,160'; $edit.Anchor='Top,Right'
$form.Controls.Add($edit)
$lblType = New-Object System.Windows.Forms.Label; $lblType.Text='Type:'; $lblType.Location='10,25'; $lblType.AutoSize=$true; $edit.Controls.Add($lblType)
$txtType = New-Object System.Windows.Forms.TextBox; $txtType.Location='60,22'; $txtType.Width=90; $txtType.ReadOnly=$true; $edit.Controls.Add($txtType)
$lblX = New-Object System.Windows.Forms.Label; $lblX.Text='XRel:'; $lblX.Location='10,55'; $lblX.AutoSize=$true; $edit.Controls.Add($lblX)
$txtX = New-Object System.Windows.Forms.TextBox; $txtX.Location='60,52'; $txtX.Width=70; $edit.Controls.Add($txtX)
$lblY = New-Object System.Windows.Forms.Label; $lblY.Text='YRel:'; $lblY.Location='140,55'; $lblY.AutoSize=$true; $edit.Controls.Add($lblY)
$txtY = New-Object System.Windows.Forms.TextBox; $txtY.Location='190,52'; $txtY.Width=70; $edit.Controls.Add($txtY)
$lblIn = New-Object System.Windows.Forms.Label; $lblIn.Text='Input (SendKeys):'; $lblIn.Location='10,85'; $lblIn.AutoSize=$true; $edit.Controls.Add($lblIn)
$txtIn = New-Object System.Windows.Forms.TextBox; $txtIn.Location='120,82'; $txtIn.Width=180; $edit.Controls.Add($txtIn)
$lblDur = New-Object System.Windows.Forms.Label; $lblDur.Text='Duration (ms):'; $lblDur.Location='10,85'; $lblDur.AutoSize=$true; $lblDur.Visible=$false; $edit.Controls.Add($lblDur)
$txtDur = New-Object System.Windows.Forms.TextBox; $txtDur.Location='120,82'; $txtDur.Width=100; $txtDur.Visible=$false; $edit.Controls.Add($txtDur)

$btnUpdate = New-Object System.Windows.Forms.Button; $btnUpdate.Text='Update'; $btnUpdate.Location='10,120'; $btnUpdate.Width=90; $edit.Controls.Add($btnUpdate)
$btnDelete = New-Object System.Windows.Forms.Button; $btnDelete.Text='Delete'; $btnDelete.Location='110,120'; $btnDelete.Width=90; $edit.Controls.Add($btnDelete)
$btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text='Cancel'; $btnCancel.Location='210,120'; $btnCancel.Width=90; $edit.Controls.Add($btnCancel)

function Populate-Editor{
  $idx = $list.SelectedIndex
  if($idx -lt 0){ $txtType.Text=''; $txtX.Text=''; $txtY.Text=''; $txtIn.Text=''; $txtDur.Text=''; return }
  $a = $global:Actions[$idx]
  $txtType.Text = $a.Type
  if($a.Type -eq 'Click'){
    $lblX.Visible=$true; $txtX.Visible=$true; $lblY.Visible=$true; $txtY.Visible=$true; $lblIn.Visible=$true; $txtIn.Visible=$true; $lblDur.Visible=$false; $txtDur.Visible=$false
    $txtX.Text = [string]$a.XRel
    $txtY.Text = [string]$a.YRel
    $txtIn.Text = if($a.Input){ [string]$a.Input } else { '' }
  } else {
    $lblX.Visible=$false; $txtX.Visible=$false; $lblY.Visible=$false; $txtY.Visible=$false; $lblIn.Visible=$false; $txtIn.Visible=$false; $lblDur.Visible=$true; $txtDur.Visible=$true
    $txtDur.Text = [string]$a.Duration
  }
}

$list.Add_SelectedIndexChanged({ Populate-Editor })

$btnUpdate.Add_Click({
  $idx = $list.SelectedIndex; if($idx -lt 0){ return }
  $a = $global:Actions[$idx]
  if($a.Type -eq 'Click'){
    try{
      $x=[double]$txtX.Text; $y=[double]$txtY.Text
      if(($x -lt 0) -or ($x -gt 1) -or ($y -lt 0) -or ($y -gt 1)){ throw 'XRel/YRel must be between 0 and 1' }
      $a.XRel = [Math]::Round($x,4); $a.YRel=[Math]::Round($y,4)
      if([string]::IsNullOrEmpty($txtIn.Text)){ $a.Input = '' } else { $a.Input = $txtIn.Text }
    } catch { [System.Windows.Forms.MessageBox]::Show("Invalid click values: " + $_.Exception.Message); return }
  } else {
    try{ $a.Duration = [int]$txtDur.Text; if($a.Duration -lt 0){ throw 'Duration must be >= 0' } }
    catch { [System.Windows.Forms.MessageBox]::Show("Invalid duration: " + $_.Exception.Message); return }
  }
  Refresh-List; $list.SelectedIndex=$idx; Populate-Editor
})

$btnDelete.Add_Click({ $idx=$list.SelectedIndex; if($idx -ge 0){ $global:Actions.RemoveAt($idx); Refresh-List; if($global:Actions.Count -gt 0){ $list.SelectedIndex=[Math]::Min($idx, $global:Actions.Count-1) } else { $list.ClearSelected() }; Populate-Editor } })
$btnCancel.Add_Click({ Populate-Editor })

# Save/Load/Run
$grp2 = New-Object System.Windows.Forms.GroupBox; $grp2.Text='Sequence'; $grp2.Location='560,380'; $grp2.Size='320,130'; $grp2.Anchor='Top,Right'; $form.Controls.Add($grp2)
$btnRun = New-Object System.Windows.Forms.Button; $btnRun.Text='Run'; $btnRun.Location='10,25'; $btnRun.Width=90; $grp2.Controls.Add($btnRun)
$btnRun.Add_Click({ Play-Sequence })
$btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text='Save…'; $btnSave.Location='110,25'; $btnSave.Width=90; $grp2.Controls.Add($btnSave)
$btnSave.Add_Click({ $dlg = New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter='JSON (*.json)|*.json|All files (*.*)|*.*'; if($dlg.ShowDialog() -eq 'OK'){ Save-Sequence $dlg.FileName } })
$btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text='Load…'; $btnLoad.Location='210,25'; $btnLoad.Width=90; $grp2.Controls.Add($btnLoad)
$btnLoad.Add_Click({ $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter='JSON (*.json)|*.json|All files (*.*)|*.*'; if($dlg.ShowDialog() -eq 'OK'){ Load-Sequence $dlg.FileName; Refresh-List } }) 
$lblDbc = New-Object System.Windows.Forms.Label; $lblDbc.Text='Before click (ms)'; $lblDbc.Location='10,60'; $lblDbc.AutoSize=$true; $grp2.Controls.Add($lblDbc)
$txtDbc = New-Object System.Windows.Forms.TextBox; $txtDbc.Location='120,57'; $txtDbc.Width=60; $txtDbc.Text=$global:DelayBeforeClick; $grp2.Controls.Add($txtDbc)
$lblDac = New-Object System.Windows.Forms.Label; $lblDac.Text='After click (ms)'; $lblDac.Location='10,85'; $lblDac.AutoSize=$true; $grp2.Controls.Add($lblDac)
$txtDac = New-Object System.Windows.Forms.TextBox; $txtDac.Location='120,82'; $txtDac.Width=60; $txtDac.Text=$global:DelayAfterClick; $grp2.Controls.Add($txtDac)
$lblDba = New-Object System.Windows.Forms.Label; $lblDba.Text='Between actions (ms)'; $lblDba.Location='190,60'; $lblDba.AutoSize=$true; $grp2.Controls.Add($lblDba)
$txtDba = New-Object System.Windows.Forms.TextBox; $txtDba.Location='300,57'; $txtDba.Width=60; $txtDba.Text=$global:DelayBetweenActions; $grp2.Controls.Add($txtDba)

foreach($tb in @($txtDbc,$txtDac,$txtDba)){
  $tb.Add_Leave({
    try{
      $global:DelayBeforeClick    = [int]$txtDbc.Text
      $global:DelayAfterClick     = [int]$txtDac.Text
      $global:DelayBetweenActions = [int]$txtDba.Text
    } catch {}
  })
}

# Init
Refresh-WindowList
Refresh-List
[void]$form.ShowDialog()

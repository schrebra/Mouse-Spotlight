[System.Windows.Forms.Application]::EnableVisualStyles()
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsFormsIntegration

# Updated config path for compiled exe support
$global:configPath = Join-Path $PSScriptRoot "MouseSpotlight.config"
if (-not $PSScriptRoot) {
    $global:configPath = Join-Path (Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)) "MouseSpotlight.config"
}

# Import required Windows API functions with corrected GetWindowLong and extended functionality
$code = @"
using System;
using System.Runtime.InteropServices;

public class WindowsAPI {
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    
    [DllImport("user32.dll", EntryPoint="GetWindowLongPtr")]
    public static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", EntryPoint="GetWindowLong")]
    public static extern int GetWindowLongInt32(IntPtr hWnd, int nIndex);
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_LAYERED = 0x80000;
    public const int WS_EX_TRANSPARENT = 0x20;
    public const int WS_EX_NOACTIVATE = 0x08000000;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public static readonly IntPtr HWND_TOP = new IntPtr(0);
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_SHOWWINDOW = 0x0040;

    public static int GetWindowLong(IntPtr hWnd, int nIndex)
    {
        if (IntPtr.Size == 8)
            return (int)GetWindowLongPtr(hWnd, nIndex);
        else
            return GetWindowLongInt32(hWnd, nIndex);
    }
}
"@

Add-Type -TypeDefinition $code -Language CSharp

# Updated Save Configuration function with error handling and JSON support
function Save-CircleConfiguration {
    try {
        $config = @{
            Color = $colorComboBox.SelectedItem
            Size = $sizeNumeric.Value
            FillTransparency = $transparencyNumeric.Value
            OutlineEnabled = $outlineToggle.Checked
            OutlineWidth = $outlineWidthNumeric.Value
            OutlineTransparency = $outlineTransNumeric.Value
        }
        
        $configJson = $config | ConvertTo-Json
        [System.IO.File]::WriteAllText($global:configPath, $configJson)
        
        [System.Windows.Forms.MessageBox]::Show(
            "Configuration saved successfully", 
            "Mouse Spotlight", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save configuration: $_", 
            "Mouse Spotlight", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Updated Load Configuration function with error handling and JSON support
function Load-CircleConfiguration {
    try {
        if (Test-Path $global:configPath) {
            $configJson = [System.IO.File]::ReadAllText($global:configPath)
            $config = $configJson | ConvertFrom-Json
            
            $colorComboBox.SelectedItem = $config.Color
            $sizeNumeric.Value = [decimal]$config.Size
            $transparencyNumeric.Value = [decimal]$config.FillTransparency
            $outlineToggle.Checked = [bool]$config.OutlineEnabled
            $outlineWidthNumeric.Value = [decimal]$config.OutlineWidth
            $outlineTransNumeric.Value = [decimal]$config.OutlineTransparency
            
            Update-CircleAppearance
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load configuration: $_", 
            "Mouse Spotlight", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to update circle appearance
function Update-CircleAppearance {
    $selectedColor = [System.Drawing.Color]::FromName($colorComboBox.SelectedItem)
    $fillOpacity = [int]$transparencyNumeric.Value
    $outlineOpacity = [int]$outlineTransNumeric.Value
    
    $global:circleShape.Fill = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromArgb($fillOpacity, $selectedColor.R, $selectedColor.G, $selectedColor.B)
    )
    
    $global:circleShape.Stroke = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.Color]::FromArgb($outlineOpacity, $selectedColor.R, $selectedColor.G, $selectedColor.B)
    )
    
    $global:circleWindow.Width = $sizeNumeric.Value * 2
    $global:circleWindow.Height = $sizeNumeric.Value * 2
    
    if ($outlineToggle.Checked) {
        $global:circleShape.StrokeThickness = $outlineWidthNumeric.Value
    } else {
        $global:circleShape.StrokeThickness = 0
    }
}

$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Transparent Circle" 
    AllowsTransparency="True"
    WindowStyle="None"
    Background="Transparent"
    Topmost="True"
    ShowInTaskbar="False"
    IsHitTestVisible="False">
    <Grid Background="Transparent">
        <Ellipse Name="CircleShape" Stroke="#FFFF0000" StrokeThickness="2" Fill="#80FF0000"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$global:circleWindow = [System.Windows.Markup.XamlReader]::Load($reader)
$global:circleShape = $global:circleWindow.FindName('CircleShape')
$global:circleWindow.Width = 100
$global:circleWindow.Height = 100

# Enhanced window initialization with additional flags for overlay support
$global:circleWindow.Add_SourceInitialized({
    $helper = New-Object System.Windows.Interop.WindowInteropHelper($global:circleWindow)
    $style = [WindowsAPI]::GetWindowLong($helper.Handle, [WindowsAPI]::GWL_EXSTYLE)
    $style = $style -bor [WindowsAPI]::WS_EX_TRANSPARENT -bor [WindowsAPI]::WS_EX_LAYERED -bor [WindowsAPI]::WS_EX_NOACTIVATE -bor [WindowsAPI]::WS_EX_TOOLWINDOW
    [void][WindowsAPI]::SetWindowLong($helper.Handle, [WindowsAPI]::GWL_EXSTYLE, $style)
    
    # Set window to be always on top with enhanced flags
    [void][WindowsAPI]::SetWindowPos(
        $helper.Handle,
        [WindowsAPI]::HWND_TOPMOST,
        0, 0, 0, 0,
        ([WindowsAPI]::SWP_NOMOVE -bor [WindowsAPI]::SWP_NOSIZE -bor [WindowsAPI]::SWP_NOACTIVATE -bor [WindowsAPI]::SWP_SHOWWINDOW)
    )
})

$global:controlForm = New-Object System.Windows.Forms.Form
$global:controlForm.Text = " Mouse Spotlight"
$global:controlForm.Width = 320
$global:controlForm.Height = 440
$global:controlForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$global:controlForm.TopMost = $false
$global:controlForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$global:controlForm.BackColor = [System.Drawing.Color]::WhiteSmoke
$global:controlForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$global:controlForm.MaximizeBox = $false

# Menu Strip with updated styling
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = [System.Drawing.Color]::WhiteSmoke
$menuStrip.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(3, 2, 0, 2)

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$fileMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$saveMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveMenuItem.Text = "Save Configuration"
$saveMenuItem.Add_Click({ Save-CircleConfiguration })

$loadMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$loadMenuItem.Text = "Load Configuration"
$loadMenuItem.Add_Click({ Load-CircleConfiguration })

$fileMenu.DropDownItems.Add($saveMenuItem)
$fileMenu.DropDownItems.Add($loadMenuItem)
$menuStrip.Items.Add($fileMenu)
$global:controlForm.Controls.Add($menuStrip)

# Main Controls Panel
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Location = New-Object System.Drawing.Point(10, 30)
$mainPanel.Size = New-Object System.Drawing.Size(290, 140)
$mainPanel.BackColor = [System.Drawing.Color]::White
$global:controlForm.Controls.Add($mainPanel)

# Color Controls with padding and alignment
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = "Circle Color:"
$colorLabel.Location = New-Object System.Drawing.Point(15, 15)
$colorLabel.AutoSize = $true
$mainPanel.Controls.Add($colorLabel)

$colorComboBox = New-Object System.Windows.Forms.ComboBox
$colorComboBox.Location = New-Object System.Drawing.Point(140, 12)
$colorComboBox.Width = 130
$colorComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[System.Drawing.KnownColor].GetFields() | Where-Object { $_.IsStatic } | ForEach-Object { $colorComboBox.Items.Add($_.Name) }
$colorComboBox.SelectedItem = "Red"
$mainPanel.Controls.Add($colorComboBox)

# Size Controls
$sizeLabel = New-Object System.Windows.Forms.Label
$sizeLabel.Text = "Circle Size:"
$sizeLabel.Location = New-Object System.Drawing.Point(15, 50)
$sizeLabel.AutoSize = $true
$mainPanel.Controls.Add($sizeLabel)

$sizeNumeric = New-Object System.Windows.Forms.NumericUpDown
$sizeNumeric.Location = New-Object System.Drawing.Point(140, 48)
$sizeNumeric.Width = 80
$sizeNumeric.Minimum = 10
$sizeNumeric.Maximum = 200
$sizeNumeric.Value = 50
$sizeNumeric.Increment = 5
$mainPanel.Controls.Add($sizeNumeric)

# Fill Transparency Controls
$transparencyLabel = New-Object System.Windows.Forms.Label
$transparencyLabel.Text = "Fill Transparency:"
$transparencyLabel.Location = New-Object System.Drawing.Point(15, 85)
$transparencyLabel.AutoSize = $true
$mainPanel.Controls.Add($transparencyLabel)

$transparencyNumeric = New-Object System.Windows.Forms.NumericUpDown
$transparencyNumeric.Location = New-Object System.Drawing.Point(140, 83)
$transparencyNumeric.Width = 80
$transparencyNumeric.Minimum = 0
$transparencyNumeric.Maximum = 255
$transparencyNumeric.Value = 128
$transparencyNumeric.Increment = 5
$mainPanel.Controls.Add($transparencyNumeric)

# Outline Controls Group Box
$outlineGroupBox = New-Object System.Windows.Forms.GroupBox
$outlineGroupBox.Text = "Outline Controls"
$outlineGroupBox.Location = New-Object System.Drawing.Point(10, 180)
$outlineGroupBox.Size = New-Object System.Drawing.Size(290, 150)
$outlineGroupBox.BackColor = [System.Drawing.Color]::White
$global:controlForm.Controls.Add($outlineGroupBox)

# Outline Toggle with improved spacing
$outlineToggle = New-Object System.Windows.Forms.CheckBox
$outlineToggle.Text = "Show Outline"
$outlineToggle.Location = New-Object System.Drawing.Point(15, 25)
$outlineToggle.AutoSize = $true
$outlineToggle.Checked = $true
$outlineGroupBox.Controls.Add($outlineToggle)

# Outline Width Controls
$outlineWidthLabel = New-Object System.Windows.Forms.Label
$outlineWidthLabel.Text = "Outline Width:"
$outlineWidthLabel.Location = New-Object System.Drawing.Point(15, 60)
$outlineWidthLabel.AutoSize = $true
$outlineGroupBox.Controls.Add($outlineWidthLabel)

$outlineWidthNumeric = New-Object System.Windows.Forms.NumericUpDown
$outlineWidthNumeric.Location = New-Object System.Drawing.Point(140, 58)
$outlineWidthNumeric.Width = 80
$outlineWidthNumeric.Minimum = 1
$outlineWidthNumeric.Maximum = 20
$outlineWidthNumeric.Value = 2
$outlineWidthNumeric.Increment = 1
$outlineGroupBox.Controls.Add($outlineWidthNumeric)

# Outline Transparency Controls
$outlineTransLabel = New-Object System.Windows.Forms.Label
$outlineTransLabel.Text = "Outline Transparency:"
$outlineTransLabel.Location = New-Object System.Drawing.Point(15, 95)
$outlineTransLabel.AutoSize = $true
$outlineGroupBox.Controls.Add($outlineTransLabel)

$outlineTransNumeric = New-Object System.Windows.Forms.NumericUpDown
$outlineTransNumeric.Location = New-Object System.Drawing.Point(140, 93)
$outlineTransNumeric.Width = 80
$outlineTransNumeric.Minimum = 0
$outlineTransNumeric.Maximum = 255
$outlineTransNumeric.Value = 255
$outlineTransNumeric.Increment = 5
$outlineGroupBox.Controls.Add($outlineTransNumeric)

# Button Panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(10, 340)
$buttonPanel.Size = New-Object System.Drawing.Size(290, 35)
$buttonPanel.BackColor = [System.Drawing.Color]::White
$global:controlForm.Controls.Add($buttonPanel)

# Updated Buttons with modern styling
$toggleButton = New-Object System.Windows.Forms.Button
$toggleButton.Text = "Toggle Circle"
$toggleButton.Location = New-Object System.Drawing.Point(15, 5)
$toggleButton.Width = 120
$toggleButton.Height = 25
$toggleButton.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$buttonPanel.Controls.Add($toggleButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Location = New-Object System.Drawing.Point(155, 5)
$exitButton.Width = 120
$exitButton.Height = 25
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$buttonPanel.Controls.Add($exitButton)

# Hotkey Label with updated styling
$hotkeyLabel = New-Object System.Windows.Forms.Label
$hotkeyLabel.Text = "Press ESC to exit"
$hotkeyLabel.Location = New-Object System.Drawing.Point(12, 385)
$hotkeyLabel.AutoSize = $true
$hotkeyLabel.ForeColor = [System.Drawing.Color]::Gray
$global:controlForm.Controls.Add($hotkeyLabel)

# Enhanced timer for circle movement with improved z-order management
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1
$timer.Add_Tick({
    $cursorPos = [System.Windows.Forms.Cursor]::Position
    $global:circleWindow.Left = $cursorPos.X - $global:circleWindow.Width / 2
    $global:circleWindow.Top = $cursorPos.Y - $global:circleWindow.Height / 2
    
    # Continuously ensure the window stays on top
    $helper = New-Object System.Windows.Interop.WindowInteropHelper($global:circleWindow)
    [void][WindowsAPI]::SetWindowPos(
        $helper.Handle,
        [WindowsAPI]::HWND_TOPMOST,
        0, 0, 0, 0,
        ([WindowsAPI]::SWP_NOMOVE -bor [WindowsAPI]::SWP_NOSIZE -bor [WindowsAPI]::SWP_NOACTIVATE)
    )
})
$timer.Start()

# Event Handlers
$colorComboBox.Add_SelectedIndexChanged({ Update-CircleAppearance })
$sizeNumeric.Add_ValueChanged({ Update-CircleAppearance })
$transparencyNumeric.Add_ValueChanged({ Update-CircleAppearance })
$outlineToggle.Add_CheckedChanged({ Update-CircleAppearance })
$outlineWidthNumeric.Add_ValueChanged({ Update-CircleAppearance })
$outlineTransNumeric.Add_ValueChanged({ Update-CircleAppearance })

$toggleButton.Add_Click({
    $global:isVisible = !$global:isVisible
    if ($global:isVisible) {
        $global:circleWindow.Show()
        $toggleButton.Text = "Hide Circle"
    } else {
        $global:circleWindow.Hide()
        $toggleButton.Text = "Show Circle"
    }
})

$exitButton.Add_Click({
    $timer.Stop()
    $global:circleWindow.Close()
    $global:controlForm.Close()
})

$global:controlForm.KeyPreview = $true
$global:controlForm.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Escape') {
        $timer.Stop()
        $global:circleWindow.Close()
        $global:controlForm.Close()
    }
})

# Add window activation prevention
$global:controlForm.Add_Activated({
    $helper = New-Object System.Windows.Interop.WindowInteropHelper($global:circleWindow)
    [void][WindowsAPI]::SetWindowPos(
        $helper.Handle,
        [WindowsAPI]::HWND_TOPMOST,
        0, 0, 0, 0,
        ([WindowsAPI]::SWP_NOMOVE -bor [WindowsAPI]::SWP_NOSIZE -bor [WindowsAPI]::SWP_NOACTIVATE -bor [WindowsAPI]::SWP_SHOWWINDOW)
    )
})

# Initialize circle visibility
$global:isVisible = $true

# Try to load configuration if it exists
if (Test-Path $global:configPath) {
    Load-CircleConfiguration
}

# Show the windows
$global:circleWindow.Show()
$global:controlForm.ShowDialog()
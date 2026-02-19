# Auto-relaunch in PowerShell 7 if running in 5.x
if ($PSVersionTable.PSVersion.Major -lt 7) {
    $pwsh = "$env:ProgramFiles\PowerShell\7\pwsh.exe"
    if (Test-Path $pwsh) {
        & $pwsh -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
        exit
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Pre-load System.Drawing.Common for PS7 / .NET Core compatibility
try { Add-Type -AssemblyName System.Drawing.Common -ErrorAction Stop } catch {}

$refAsms = [System.AppDomain]::CurrentDomain.GetAssemblies() |
Where-Object { $_.GetName().Name -in @('System.Drawing.Common', 'System.Drawing.Primitives') } |
ForEach-Object { $_.Location } | Where-Object { $_ }

if (-not ([System.Management.Automation.PSTypeName]'ScriptLauncherHelper').Type) {
    try {
        Add-Type -IgnoreWarnings -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Drawing2D;
public class ScriptLauncherHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("dwmapi.dll")]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    public static GraphicsPath RoundedRect(Rectangle bounds, int radius) {
        var path = new GraphicsPath();
        int d = radius * 2;
        path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
        path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
        path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }
}
"@ -ReferencedAssemblies $refAsms
    }
    catch {}
}

[ScriptLauncherHelper]::SetProcessDPIAware() | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

# --- Colors ---
$colBg = [System.Drawing.Color]::FromArgb(22, 22, 26)
$colTitleBar = [System.Drawing.Color]::FromArgb(16, 16, 19)
$colAccent = [System.Drawing.Color]::FromArgb(56, 132, 244)
$colBtnDefault = [System.Drawing.Color]::FromArgb(40, 40, 46)
$colBtnHover = [System.Drawing.Color]::FromArgb(55, 55, 64)
$colBtnPress = [System.Drawing.Color]::FromArgb(34, 34, 40)
$colGreen = [System.Drawing.Color]::FromArgb(46, 204, 113)
$colRed = [System.Drawing.Color]::FromArgb(231, 76, 60)
$colYellow = [System.Drawing.Color]::FromArgb(241, 196, 15)
$colGray = [System.Drawing.Color]::FromArgb(120, 120, 135)

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Jordan's Script Launcher"
$form.Size = New-Object System.Drawing.Size(960, 700)
$form.MinimumSize = New-Object System.Drawing.Size(620, 400)
$form.StartPosition = "CenterScreen"
$form.BackColor = $colBg
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.GetType().GetProperty("DoubleBuffered",
    [System.Reflection.BindingFlags]"Instance,NonPublic").SetValue($form, $true, $null)

try {
    $pref = 2
    [ScriptLauncherHelper]::DwmSetWindowAttribute(
        $form.Handle, 33, [ref]$pref, [System.Runtime.InteropServices.Marshal]::SizeOf($pref)
    ) | Out-Null
}
catch {}

# --- System Tray ---
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Text = "Jordan's Script Launcher"
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Visible = $false

# Tray context menu
$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayShow = $trayMenu.Items.Add("Show")
$trayShow.Add_Click({
        $form.Show()
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $form.Activate()
        $trayIcon.Visible = $false
    })
$trayExit = $trayMenu.Items.Add("Exit")
$trayExit.Add_Click({ $form.Close() })
$trayIcon.ContextMenuStrip = $trayMenu

# Double-click tray icon to restore
$trayIcon.Add_DoubleClick({
        $form.Show()
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $form.Activate()
        $trayIcon.Visible = $false
    })

# Minimize to tray instead of taskbar
$form.Add_Resize({
        if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) {
            $form.Hide()
            $trayIcon.Visible = $true
            $trayIcon.ShowBalloonTip(1500, "Jordan's Script Launcher", "Minimized to system tray", [System.Windows.Forms.ToolTipIcon]::Info)
        }
    })

# --- Title Bar ---
$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Dock = [System.Windows.Forms.DockStyle]::Top
$titlePanel.Height = 56
$titlePanel.BackColor = $colTitleBar

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = [char]0x26A1 + "  Jordan's Script Launcher"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $colAccent
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$titleLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$titleLabel.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
$refreshBtn = New-Object System.Windows.Forms.Button
$refreshBtn.Text = "↻"  # Refresh symbol
$refreshBtn.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$refreshBtn.ForeColor = $colGreen
$refreshBtn.BackColor = [System.Drawing.Color]::Transparent
$refreshBtn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshBtn.FlatAppearance.BorderSize = 0
$refreshBtn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(40, 255, 255, 255)
$refreshBtn.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(20, 0, 0, 0)
$refreshBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$refreshBtn.Size = New-Object System.Drawing.Size(40, 56)
$refreshBtn.Dock = [System.Windows.Forms.DockStyle]::Right
$refreshBtn.Add_Click({ Reload-Scripts })

$titlePanel.Controls.Add($refreshBtn)
$titlePanel.Controls.Add($titleLabel)
$form.Controls.Add($titlePanel)

# --- Button Area ---
$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$buttonPanel.WrapContents = $true
$buttonPanel.Location = New-Object System.Drawing.Point(0, 56)
$buttonPanel.Anchor = "Top,Bottom,Left,Right"
$buttonPanel.AutoScroll = $true
$buttonPanel.BackColor = $colBg
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 8)
$form.Controls.Add($buttonPanel)

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
$statusBar.Height = 26
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(26, 26, 30)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.ForeColor = $colGray
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.25)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusLabel.Padding = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$statusLabel.Text = "Ready"

$statusDot = New-Object System.Windows.Forms.Panel
$statusDot.Size = New-Object System.Drawing.Size(8, 8)
$statusDot.Location = New-Object System.Drawing.Point(4, 9)
$statusDot.BackColor = $colGreen
$gp = New-Object System.Drawing.Drawing2D.GraphicsPath
$gp.AddEllipse(0, 0, 7, 7)
$statusDot.Region = New-Object System.Drawing.Region($gp)

$statusBar.Controls.Add($statusDot)
$statusBar.Controls.Add($statusLabel)
$form.Controls.Add($statusBar)

# --- Button Factory ---
$btnRadius = 10

function New-ScriptButton {
    param([string]$Label, [string]$ScriptPath)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "   " + $Label
    $btn.Tag = $ScriptPath
    $btn.Size = New-Object System.Drawing.Size(0, 42)
    $btn.Anchor = "Top,Left"
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = $colBtnDefault
    $btn.ForeColor = [System.Drawing.Color]::FromArgb(210, 210, 220)
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = $colBtnHover
    $btn.FlatAppearance.MouseDownBackColor = $colBtnPress
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $btn.Padding = New-Object System.Windows.Forms.Padding(14, 0, 0, 0)
    $btn.Margin = New-Object System.Windows.Forms.Padding(5) # Spacing handled by Margin

    # Tooltip + right-click edit
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($btn, $ScriptPath)
    $ctx = New-Object System.Windows.Forms.ContextMenuStrip
    $editItem = $ctx.Items.Add("Open in VS Code")
    $editItem.Tag = $ScriptPath
    $editItem.Add_Click({ Start-Process "code" -ArgumentList $this.Tag })
    $btn.ContextMenuStrip = $ctx

    # Rounded corners
    $btn.Add_Resize({
            param($s, $e)
            if ($s.Width -gt 0 -and $s.Height -gt 0) {
                $path = [ScriptLauncherHelper]::RoundedRect(
                    (New-Object System.Drawing.Rectangle(0, 0, $s.Width, $s.Height)), $btnRadius)
                $s.Region = New-Object System.Drawing.Region($path)
            }
        })

    # Click: launch in visible console, poll with Timer
    $btn.Add_Click({
            $scriptPath = $this.Tag
            $label = $this.Text.Trim()
            foreach ($b in $allButtons) { $b.Enabled = $false }

            $statusLabel.Text = "Running: $label..."
            $statusLabel.ForeColor = $colYellow
            $statusDot.BackColor = $colYellow

            try {
                $proc = Start-Process -FilePath "pwsh" `
                    -ArgumentList "-ExecutionPolicy", "Bypass", "-NoExit", "-File", "`"$scriptPath`"" `
                    -WorkingDirectory $scriptDir -PassThru

                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 500
                $timer.Tag = @{ Process = $proc; Label = $label }
                $timer.Add_Tick({
                        $info = $this.Tag
                        if ($info.Process.HasExited) {
                            $this.Stop(); $this.Dispose()
                            foreach ($b in $allButtons) { $b.Enabled = $true }
                            if ($info.Process.ExitCode -eq 0 -or $info.Process.ExitCode -eq -1073741510) {
                                $statusLabel.Text = "Completed: $($info.Label)"
                                $statusLabel.ForeColor = $colGreen
                                $statusDot.BackColor = $colGreen
                            }
                            else {
                                $statusLabel.Text = "Error: $($info.Label) (exit $($info.Process.ExitCode))"
                                $statusLabel.ForeColor = $colRed
                                $statusDot.BackColor = $colRed
                            }
                        }
                    })
                $timer.Start()
            }
            catch {
                foreach ($b in $allButtons) { $b.Enabled = $true }
                $statusLabel.Text = "Error: $label"
                $statusLabel.ForeColor = $colRed
                $statusDot.BackColor = $colRed
            }
        })

    return $btn
}

# --- Script Discovery & Reload ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$myFullPath = (Resolve-Path $MyInvocation.MyCommand.Path).Path
$btnGap = 10
$allButtons = @()

function Reload-Scripts {
    $buttonPanel.SuspendLayout()
    
    # Clear existing buttons safely
    while ($buttonPanel.Controls.Count -gt 0) {
        $btn = $buttonPanel.Controls[0]
        $buttonPanel.Controls.Remove($btn)
        $btn.Dispose()
    }
    $script:allButtons = @()

    # Find scripts
    $otherScripts = Get-ChildItem -Path $scriptDir -Filter "*.ps1" |
    Where-Object { $_.FullName -ne $myFullPath } |
    Sort-Object @{Expression = { if ($_.BaseName -match '^\d') { 0 } else { 1 } } },
    @{Expression = { if ($_.BaseName -match '^\d+') { [int]$Matches[0] } else { 0 } } },
    Name

    # Create buttons
    foreach ($s in $otherScripts) {
        $lines = Get-Content $s.FullName -First 2
        $label = $s.Name
        foreach ($line in $lines) {
            if ($line -match '^#(?!Requires\b|!)') {
                $label = $line.TrimStart("#").Trim()
                break
            }
        }
        $btn = New-ScriptButton -Label $label -ScriptPath $s.FullName
        $buttonPanel.Controls.Add($btn)
        $script:allButtons += $btn
    }
    
    $buttonPanel.ResumeLayout()
    Update-Layout
    $statusLabel.Text = "Ready  |  $($script:allButtons.Count) script(s) loaded"
}

# --- Layout ---
function Update-Layout {
    $cw = $buttonPanel.ClientSize.Width
    $ch = $form.ClientSize.Height
    $buttonPanel.Width = $form.ClientSize.Width
    $buttonPanel.Height = $ch - 56 - 26  # title - status
    
    # Calculate button width: (ContainerWidth - Padding - Scrollbar) / 2 - Margins
    # Padding=32 (16*2), Margins=20 (5*4), Scrollbar safety=12
    $btnW = [Math]::Floor(($cw - 64) / 2)
    if ($btnW -lt 200) { $btnW = $cw - 42 } # Fallback to 1 column

    foreach ($btn in $allButtons) {
        if ($btn.Visible) {
            if ($btn.Width -ne $btnW) { $btn.Width = $btnW }
        }
    }
}

$form.Add_Resize({ Update-Layout })
$form.Add_Shown({
        Reload-Scripts
    
        $btnRows = [Math]::Ceiling($allButtons.Count / 2)
        $idealClient = 56 + ($btnRows * 50) + 16 + 26
        $chrome = $form.Height - $form.ClientSize.Height
        $ideal = [Math]::Min($idealClient + $chrome, [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height)
        $form.Height = [Math]::Max($ideal, $form.MinimumSize.Height)
    
        Update-Layout
    })

$form.Add_FormClosing({
        $trayIcon.Visible = $false
        $trayIcon.Dispose()
    })

[System.Windows.Forms.Application]::Run($form)

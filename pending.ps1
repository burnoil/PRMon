Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Function to check registry key and value existence with truncation
function Check-RegistryValue {
    param (
        [string]$Path,
        [string]$ValueName
    )
    try {
        $value = Get-ItemProperty -Path $Path -Name $ValueName -ErrorAction Stop
        $valueString = "$($value.$ValueName)"
        if ($valueString.Length -gt 100) {
            $valueString = $valueString.Substring(0, 97) + "..."
        }
        if ($ValueName -eq "UpdateExeVolatile") {
            return [PSCustomObject]@{
                Result = "Found: $Path\$ValueName = $valueString"
                IndicatesRestart = ($value.$ValueName -ne 0)
            }
        }
        if ($value.$ValueName) {
            return [PSCustomObject]@{
                Result = "Found: $Path\$ValueName = $valueString"
                IndicatesRestart = $true
            }
        }
        return [PSCustomObject]@{
            Result = "Found: $Path\$ValueName = $valueString (no restart indicated)"
            IndicatesRestart = $false
        }
    }
    catch {
        return [PSCustomObject]@{
            Result = "Not found: $Path\$ValueName"
            IndicatesRestart = $false
        }
    }
}

# Function to check for subkeys (custom logic for Services\Pending)
function Check-RegistrySubkeys {
    param (
        [string]$Path
    )
    try {
        if (Test-Path $Path) {
            $subkeys = Get-ChildItem -Path $Path -ErrorAction Stop
            if ($subkeys.Count -gt 0) {
                return [PSCustomObject]@{
                    Result = "Found: $Path has $($subkeys.Count) pending subkeys"
                    IndicatesRestart = $true
                }
            }
            if ($Path -eq "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending") {
                return [PSCustomObject]@{
                    Result = "Not found: $Path (no subkeys)"
                    IndicatesRestart = $false
                }
            }
            return [PSCustomObject]@{
                Result = "Found: $Path exists (no subkeys)"
                IndicatesRestart = $false
            }
        }
        return [PSCustomObject]@{
            Result = "Not found: $Path"
            IndicatesRestart = $false
        }
    }
    catch {
        return [PSCustomObject]@{
            Result = "Not found: $Path (access error)"
            IndicatesRestart = $false
        }
    }
}

# Registry locations to check (values, SCCM paths removed)
$registryChecks = @(
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"; ValueName = "RebootRequired" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"; ValueName = "RebootPending" },
    @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"; ValueName = "PendingFileRenameOperations" },
    @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"; ValueName = "PendingFileRenameOperations2" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Updates"; ValueName = "UpdateExeVolatile" },
    @{ Path = "HKLM:\SOFTWARE\BigFix\EnterpriseClient\BESPendingRestart"; ValueName = "BESPendingRestart" },
    @{ Path = "HKLM:\SOFTWARE\Wow6432Node\BigFix\EnterpriseClient\BESPendingRestart"; ValueName = "BESPendingRestart" },
    @{ Path = "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Settings\Client"; ValueName = "_BESClient_RebootPending" },
    @{ Path = "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Global"; ValueName = "RebootPending" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServicesOnce"; ValueName = "BESPendingRestart" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"; ValueName = "PostRebootReporting" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"; ValueName = "DVDRebootSignal" },
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer"; ValueName = "InProgress" },
    @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon"; ValueName = "JoinDomain" },
    @{ Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon"; ValueName = "AvoidSpnSet" }
)

# Subkey/key existence checks (SCCM paths not present)
$subkeyChecks = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending",
    "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Actions",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending",
    "HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts"
)

# Create the form with compact size
$form = New-Object System.Windows.Forms.Form
$form.Text = "Pending Restart Monitor"
$form.Size = New-Object System.Drawing.Size(800, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize = New-Object System.Drawing.Size(400, 300)

# Create RichTextBox
$richTextBox = New-Object System.Windows.Forms.RichTextBox
$richTextBox.Location = New-Object System.Drawing.Point(10, 10)
$richTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$richTextBox.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 20), ($form.ClientSize.Height - 130))
$richTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$richTextBox.WordWrap = $true
$richTextBox.ReadOnly = $true
$richTextBox.ScrollBars = "Vertical"
$form.Controls.Add($richTextBox)

# Create Refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(10, ($form.ClientSize.Height - 100))
$refreshButton.Size = New-Object System.Drawing.Size(100, 30)
$refreshButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$refreshButton.Text = "Manual Refresh"
$refreshButton.Add_Click({ Update-Display })
$form.Controls.Add($refreshButton)

# Create Export button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(120, ($form.ClientSize.Height - 100))
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$exportButton.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$exportButton.Text = "Export"
$exportButton.Add_Click({
    $filePath = [System.Windows.Forms.SaveFileDialog]::new()
    $filePath.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $filePath.FileName = "PendingRestartStatus_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($filePath.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $richTextBox.Text | Out-File -FilePath $filePath.FileName
    }
})
$form.Controls.Add($exportButton)

# Create Checkboxes for refresh rates
$checkBox10s = New-Object System.Windows.Forms.CheckBox
$checkBox10s.Location = New-Object System.Drawing.Point(10, ($form.ClientSize.Height - 65))
$checkBox10s.Size = New-Object System.Drawing.Size(70, 20)
$checkBox10s.Text = "10 sec"
$checkBox10s.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($checkBox10s)

$checkBox30s = New-Object System.Windows.Forms.CheckBox
$checkBox30s.Location = New-Object System.Drawing.Point(90, ($form.ClientSize.Height - 65))
$checkBox30s.Size = New-Object System.Drawing.Size(70, 20)
$checkBox30s.Text = "30 sec"
$checkBox30s.Checked = $true
$checkBox30s.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($checkBox30s)

$checkBox1m = New-Object System.Windows.Forms.CheckBox
$checkBox1m.Location = New-Object System.Drawing.Point(170, ($form.ClientSize.Height - 65))
$checkBox1m.Size = New-Object System.Drawing.Size(70, 20)
$checkBox1m.Text = "1 min"
$checkBox1m.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($checkBox1m)

$checkBox5m = New-Object System.Windows.Forms.CheckBox
$checkBox5m.Location = New-Object System.Drawing.Point(250, ($form.ClientSize.Height - 65))
$checkBox5m.Size = New-Object System.Drawing.Size(70, 20)
$checkBox5m.Text = "5 min"
$checkBox5m.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($checkBox5m)

# Create Pause Refresh checkbox
$checkBoxPause = New-Object System.Windows.Forms.CheckBox
$checkBoxPause.Location = New-Object System.Drawing.Point(330, ($form.ClientSize.Height - 65))
$checkBoxPause.Size = New-Object System.Drawing.Size(100, 20)
$checkBoxPause.Text = "Pause Refresh"
$checkBoxPause.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$checkBoxPause.Add_Click({
    $timer.Enabled = -not $checkBoxPause.Checked
})
$form.Controls.Add($checkBoxPause)

# Create StatusStrip
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$statusStripLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStripLabel.Text = "Last updated: Never | Refresh: 30 sec"
$statusStrip.Items.Add($statusStripLabel)
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom
$form.Controls.Add($statusStrip)

# Function to update timer interval and ensure single selection
function Update-TimerInterval {
    param (
        [int]$Interval,
        [System.Windows.Forms.CheckBox]$SelectedCheckBox
    )
    $timer.Stop()
    $timer.Interval = $Interval
    $refreshText = switch ($Interval) {
        10000 { "10 sec" }
        30000 { "30 sec" }
        60000 { "1 min" }
        300000 { "5 min" }
    }
    $statusStripLabel.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Refresh: $refreshText" + $(if ($checkBoxPause.Checked) { " (Paused)" } else { "" })
    if (-not $checkBoxPause.Checked) {
        $timer.Start()
    }
    foreach ($cb in @($checkBox10s, $checkBox30s, $checkBox1m, $checkBox5m)) {
        if ($cb -ne $SelectedCheckBox) {
            $cb.Checked = $false
        }
    }
}

# Add click events for checkboxes
$checkBox10s.Add_Click({
    if ($checkBox10s.Checked) {
        Update-TimerInterval -Interval 10000 -SelectedCheckBox $checkBox10s
    }
})
$checkBox30s.Add_Click({
    if ($checkBox30s.Checked) {
        Update-TimerInterval -Interval 30000 -SelectedCheckBox $checkBox30s
    }
})
$checkBox1m.Add_Click({
    if ($checkBox1m.Checked) {
        Update-TimerInterval -Interval 60000 -SelectedCheckBox $checkBox1m
    }
})
$checkBox5m.Add_Click({
    if ($checkBox5m.Checked) {
        Update-TimerInterval -Interval 300000 -SelectedCheckBox $checkBox5m
    }
})

# Update functions with compact visual enhancements
function Update-RegistryChecks {
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $richTextBox.SelectionColor = [System.Drawing.Color]::DarkBlue
    $richTextBox.AppendText("Checking registry locations for pending restart indicators...`n")
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("-------------------------------------------------------------`n")
    foreach ($check in $registryChecks) {
        $result = Check-RegistryValue -Path $check.Path -ValueName $check.ValueName
        $richTextBox.SelectionColor = if ($result.IndicatesRestart) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }
        $richTextBox.AppendText($result.Result + "`n")
    }
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("---`n")
    $richTextBox.Update()
}

function Update-SubkeyChecks {
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $richTextBox.SelectionColor = [System.Drawing.Color]::DarkBlue
    $richTextBox.AppendText("Checking registry locations with pending subkeys/keys...`n")
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("-------------------------------------------------------------`n")
    foreach ($path in $subkeyChecks) {
        $result = Check-RegistrySubkeys -Path $path
        $richTextBox.SelectionColor = if ($result.IndicatesRestart) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }
        $richTextBox.AppendText($result.Result + "`n")
    }
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("---`n")
    $richTextBox.Update()
}

function Update-SystemStatus {
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $richTextBox.SelectionColor = [System.Drawing.Color]::DarkBlue
    $richTextBox.AppendText("Checking system reboot status...`n")
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
    try {
        $rebootPending = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name "RebootRequired" -ErrorAction SilentlyContinue) -ne $null
        $richTextBox.SelectionColor = if ($rebootPending) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }
        $richTextBox.AppendText($(if ($rebootPending) { "System indicates a reboot is pending" } else { "No system-level reboot pending detected" }) + "`n")
    } catch {
        $richTextBox.SelectionColor = [System.Drawing.Color]::Orange
        $richTextBox.AppendText("Unable to check system reboot status`n")
    }
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("---`n")
    $richTextBox.Update()
}

function Update-BigFixStatus {
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $richTextBox.SelectionColor = [System.Drawing.Color]::DarkBlue
    $richTextBox.AppendText("Checking BigFix-specific reboot status...`n")
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
    $bigfixPending = $false
    foreach ($path in @("HKLM:\SOFTWARE\BigFix\EnterpriseClient\BESPendingRestart", "HKLM:\SOFTWARE\Wow6432Node\BigFix\EnterpriseClient\BESPendingRestart")) {
        if (Test-Path $path) {
            $value = Get-ItemProperty -Path $path -Name "BESPendingRestart" -ErrorAction SilentlyContinue
            if ($value -and $value.BESPendingRestart) {
                $bigfixPending = $true
                $richTextBox.SelectionColor = [System.Drawing.Color]::Red
                $richTextBox.AppendText("BigFix indicates a reboot is pending at $path`n")
            }
        }
    }
    if (-not $bigfixPending) {
        $richTextBox.SelectionColor = [System.Drawing.Color]::Green
        $richTextBox.AppendText("No BigFix-specific reboot pending detected`n")
    }
    $richTextBox.SelectionColor = [System.Drawing.Color]::Gray
    $richTextBox.AppendText("---`n")
    $richTextBox.Update()
}

function Update-ComputerNameStatus {
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $richTextBox.SelectionColor = [System.Drawing.Color]::DarkBlue
    $richTextBox.AppendText("Checking ComputerName status...`n")
    $richTextBox.SelectionFont = New-Object System.Drawing.Font("Consolas", 10)
    try {
        $activeName = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName" -Name "ComputerName" -ErrorAction SilentlyContinue).ComputerName
        $pendingName = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName" -Name "ComputerName" -ErrorAction SilentlyContinue).ComputerName
        if ($activeName -and $pendingName) {
            $namesMatch = $activeName -eq $pendingName
            $richTextBox.SelectionColor = if (-not $namesMatch) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Green }
            $richTextBox.AppendText("ComputerName: Active = $activeName, Pending = $pendingName" + $(if ($namesMatch) { " (match)" } else { " (do not match)" }) + "`n")
        } else {
            $richTextBox.SelectionColor = [System.Drawing.Color]::Orange
            $richTextBox.AppendText("ComputerName: Unable to retrieve one or both names`n")
        }
    } catch {
        $richTextBox.SelectionColor = [System.Drawing.Color]::Orange
        $richTextBox.AppendText("Unable to check ComputerName status`n")
    }
    $richTextBox.Update()
}

# Main update function
function Update-Display {
    $richTextBox.Clear()
    Update-RegistryChecks
    Update-SubkeyChecks
    Update-SystemStatus
    Update-BigFixStatus
    Update-ComputerNameStatus
    $refreshText = switch ($timer.Interval) {
        10000 { "10 sec" }
        30000 { "30 sec" }
        60000 { "1 min" }
        300000 { "5 min" }
    }
    $statusStripLabel.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Refresh: $refreshText" + $(if ($checkBoxPause.Checked) { " (Paused)" } else { "" })
    $form.Update()
    [System.Windows.Forms.Application]::DoEvents()
}

# Handle form resize
$form.Add_Resize({
    $richTextBox.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 20), ($form.ClientSize.Height - 130))
    $refreshButton.Location = New-Object System.Drawing.Point(10, ($form.ClientSize.Height - 100))
    $exportButton.Location = New-Object System.Drawing.Point(120, ($form.ClientSize.Height - 100))
    $checkBox10s.Location = New-Object System.Drawing.Point(10, ($form.ClientSize.Height - 65))
    $checkBox30s.Location = New-Object System.Drawing.Point(90, ($form.ClientSize.Height - 65))
    $checkBox1m.Location = New-Object System.Drawing.Point(170, ($form.ClientSize.Height - 65))
    $checkBox5m.Location = New-Object System.Drawing.Point(250, ($form.ClientSize.Height - 65))
    $checkBoxPause.Location = New-Object System.Drawing.Point(330, ($form.ClientSize.Height - 65))
})

# Create timer for auto-refresh (default 30 seconds)
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000
$timer.Add_Tick({ 
    Update-Display 
})

# Show the form and update after it’s visible
$form.Show()
$form.Activate()
Update-Display
$timer.Start()
[System.Windows.Forms.Application]::Run($form)

# Clean up
$timer.Stop()
$timer.Dispose()
$form.Dispose()
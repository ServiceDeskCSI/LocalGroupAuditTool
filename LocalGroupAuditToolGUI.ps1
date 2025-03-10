# --- Determine script folder and configuration file paths ---
if ($PSScriptRoot) { 
    $scriptDir = $PSScriptRoot 
} else { 
    $scriptDir = (Get-Location).Path 
}
$computersConfigPath = Join-Path -Path $scriptDir -ChildPath "computers.config"

# CSV file paths for successful and failed scans (for Local Group Members)
$successCSV = Join-Path -Path $scriptDir -ChildPath "LocalGroupMembers_Success.csv"
$failureCSV = Join-Path -Path $scriptDir -ChildPath "LocalGroupMembers_Failure.csv"

# --- Load required assemblies ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Load machines from config file if it exists ---
$machineList = @()
if (Test-Path $computersConfigPath) {
    $machineList = Get-Content -Path $computersConfigPath | Where-Object { $_.Trim() -ne "" }
}

# --- Create the Main Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Local Groups Audit Tool"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"

# --- Create CheckedListBox for machine list ---
$clbMachines = New-Object System.Windows.Forms.CheckedListBox
$clbMachines.Location = New-Object System.Drawing.Point(10,10)
$clbMachines.Size = New-Object System.Drawing.Size(250,350)
$clbMachines.CheckOnClick = $true
$clbMachines.BeginUpdate()
foreach ($m in $machineList) { [void]$clbMachines.Items.Add($m) }
$clbMachines.EndUpdate()
for ($i=0; $i -lt $clbMachines.Items.Count; $i++) { $clbMachines.SetItemChecked($i, $true) }
$form.Controls.Add($clbMachines)

# --- Editing Buttons for the machine list ---
$btnUp = New-Object System.Windows.Forms.Button; 
$btnUp.Text = "Up"; 
$btnUp.Size = New-Object System.Drawing.Size(60,30); 
$btnUp.Location = New-Object System.Drawing.Point(270,10); 
$form.Controls.Add($btnUp)

$btnDown = New-Object System.Windows.Forms.Button; 
$btnDown.Text = "Down"; 
$btnDown.Size = New-Object System.Drawing.Size(60,30); 
$btnDown.Location = New-Object System.Drawing.Point(270,50); 
$form.Controls.Add($btnDown)

$btnDelete = New-Object System.Windows.Forms.Button; 
$btnDelete.Text = "Delete"; 
$btnDelete.Size = New-Object System.Drawing.Size(60,30); 
$btnDelete.Location = New-Object System.Drawing.Point(270,90); 
$form.Controls.Add($btnDelete)

$btnSaveList = New-Object System.Windows.Forms.Button; 
$btnSaveList.Text = "Save List"; 
$btnSaveList.Size = New-Object System.Drawing.Size(60,30); 
$btnSaveList.Location = New-Object System.Drawing.Point(270,130); 
$form.Controls.Add($btnSaveList)

# --- Button to Get AD Computers ---
$btnGetAD = New-Object System.Windows.Forms.Button; 
$btnGetAD.Text = "Get AD Computers"; 
$btnGetAD.Size = New-Object System.Drawing.Size(120,30); 
$btnGetAD.Location = New-Object System.Drawing.Point(10,370); 
$form.Controls.Add($btnGetAD)
$btnGetAD.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adComputers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name
        if ($adComputers.Count -gt 0) {
            $clbMachines.BeginUpdate()
            $clbMachines.Items.Clear()
            foreach ($comp in $adComputers) { [void]$clbMachines.Items.Add($comp) }
            $clbMachines.EndUpdate()
            for ($i=0; $i -lt $clbMachines.Items.Count; $i++) { $clbMachines.SetItemChecked($i, $true) }
            $adComputers | Out-File -FilePath $computersConfigPath -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Retrieved $($adComputers.Count) computers from AD.", "Get AD Computers")
        } else {
            [System.Windows.Forms.MessageBox]::Show("No computers found in AD.", "Get AD Computers")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error retrieving AD computers: " + $_.Exception.Message, "Error")
    }
})

# --- Audit Options ---
$lblLocalGroup = New-Object System.Windows.Forms.Label; 
$lblLocalGroup.Text = "Local Group:"; 
$lblLocalGroup.AutoSize = $true; 
$lblLocalGroup.Location = New-Object System.Drawing.Point(400,10); 
$form.Controls.Add($lblLocalGroup)
$txtLocalGroup = New-Object System.Windows.Forms.TextBox; 
$txtLocalGroup.Text = "Administrators"; 
$txtLocalGroup.Location = New-Object System.Drawing.Point(500,10); 
$txtLocalGroup.Size = New-Object System.Drawing.Size(100,20); 
$form.Controls.Add($txtLocalGroup)

$lblRescan = New-Object System.Windows.Forms.Label; 
$lblRescan.Text = "Rescan Interval (hours):"; 
$lblRescan.AutoSize = $true; 
$lblRescan.Location = New-Object System.Drawing.Point(400,40); 
$form.Controls.Add($lblRescan)
$txtRescan = New-Object System.Windows.Forms.TextBox; 
$txtRescan.Text = "2"; 
$txtRescan.Location = New-Object System.Drawing.Point(550,40); 
$txtRescan.Size = New-Object System.Drawing.Size(40,20); 
$form.Controls.Add($txtRescan)

$btnRunAudit = New-Object System.Windows.Forms.Button; 
$btnRunAudit.Text = "Run Audit"; 
$btnRunAudit.Size = New-Object System.Drawing.Size(120,30); 
$btnRunAudit.Location = New-Object System.Drawing.Point(400,80); 
$form.Controls.Add($btnRunAudit)

$btnStopScan = New-Object System.Windows.Forms.Button; 
$btnStopScan.Text = "Stop Scan"; 
$btnStopScan.Size = New-Object System.Drawing.Size(120,30); 
$btnStopScan.Location = New-Object System.Drawing.Point(530,80); 
$form.Controls.Add($btnStopScan)

# --- Add a small console window (TextBox) under the scan buttons ---
$consoleTextBox = New-Object System.Windows.Forms.TextBox
$consoleTextBox.Location = New-Object System.Drawing.Point(400,120)
$consoleTextBox.Size = New-Object System.Drawing.Size(320,100)
$consoleTextBox.Multiline = $true
$consoleTextBox.ScrollBars = 'Vertical'
$consoleTextBox.ReadOnly = $true
$consoleTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
$form.Controls.Add($consoleTextBox)

# --- Progress Bar and Status Label ---
$progressBar = New-Object System.Windows.Forms.ProgressBar; 
$progressBar.Location = New-Object System.Drawing.Point(10,530); 
$progressBar.Size = New-Object System.Drawing.Size(760,20); 
$progressBar.Minimum = 0; 
$progressBar.Value = 0; 
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label; 
$statusLabel.AutoSize = $true; 
$statusLabel.Location = New-Object System.Drawing.Point(10,560); 
$statusLabel.Text = "Ready."; 
$form.Controls.Add($statusLabel)

# --- Global Variables ---
$global:StopScan = $false

# --- Helper function to write to the console window ---
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $consoleTextBox.AppendText("[$timestamp] $message`r`n")
    [System.Windows.Forms.Application]::DoEvents()
}

# --- Helper function to append a line to a CSV file ---
function Append-ToCSV {
    param(
        [string]$FilePath,
        [string]$Line
    )
    $Line | Out-File -FilePath $FilePath -Append -Encoding UTF8
}

# --- Function to save the current machine list to computers.config ---
function Save-MachineList {
    $list = @()
    foreach ($item in $clbMachines.Items) { $list += $item.ToString() }
    $list | Out-File -FilePath $computersConfigPath -Encoding UTF8
}

# --- Button Event Handlers for Editing the Machine List ---
$btnUp.Add_Click({
    $index = $clbMachines.SelectedIndex
    if ($index -gt 0) {
        $item = $clbMachines.Items[$index]
        $checked = $clbMachines.GetItemChecked($index)
        $clbMachines.Items.RemoveAt($index)
        $clbMachines.Items.Insert($index-1, $item)
        $clbMachines.SetItemChecked($index-1, $checked)
        $clbMachines.SelectedIndex = $index-1
    }
})
$btnDown.Add_Click({
    $index = $clbMachines.SelectedIndex
    if ($index -ge 0 -and $index -lt $clbMachines.Items.Count - 1) {
        $item = $clbMachines.Items[$index]
        $checked = $clbMachines.GetItemChecked($index)
        $clbMachines.Items.RemoveAt($index)
        $clbMachines.Items.Insert($index+1, $item)
        $clbMachines.SetItemChecked($index+1, $checked)
        $clbMachines.SelectedIndex = $index+1
    }
})
$btnDelete.Add_Click({
    $index = $clbMachines.SelectedIndex
    if ($index -ge 0) {
        $clbMachines.Items.RemoveAt($index)
    }
})
$btnSaveList.Add_Click({
    Save-MachineList
    [System.Windows.Forms.MessageBox]::Show("Machine list saved to $computersConfigPath", "Save List")
})

$btnGetAD.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adComputers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name
        if ($adComputers.Count -gt 0) {
            $clbMachines.BeginUpdate()
            $clbMachines.Items.Clear()
            foreach ($comp in $adComputers) { [void]$clbMachines.Items.Add($comp) }
            $clbMachines.EndUpdate()
            for ($i=0; $i -lt $clbMachines.Items.Count; $i++) { $clbMachines.SetItemChecked($i, $true) }
            Save-MachineList
            Write-Log "Retrieved $($adComputers.Count) computers from AD."
            [System.Windows.Forms.MessageBox]::Show("Retrieved $($adComputers.Count) computers from AD.", "Get AD Computers")
        } else {
            Write-Log "No computers found in AD."
            [System.Windows.Forms.MessageBox]::Show("No computers found in AD.", "Get AD Computers")
        }
    } catch {
        Write-Log "Error retrieving AD computers: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error retrieving AD computers: " + $_.Exception.Message, "Error")
    }
})

# --- Function for scanning a computer using ADSI ---
function Scan-Computer {
    param(
        [string]$Computer,
        [string]$LocalGroupName
    )
    Write-Log "Working on $Computer"
    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) {
        Write-Log "$Computer is offline."
        return @{ Computer = $Computer; LocalGroupName = $LocalGroupName; Status = "Offline"; Members = $null }
    }
    try {
        $group = [ADSI]"WinNT://$Computer/$LocalGroupName"
        $members = @($group.Invoke("Members"))
        if (-not $members) {
            Write-Log "No members found in group on $Computer."
            return @{ Computer = $Computer; LocalGroupName = $LocalGroupName; Status = "NoMembersFound"; Members = $null }
        }
    }
    catch {
        Write-Log "Failed to query the members of $Computer."
        return @{ Computer = $Computer; LocalGroupName = $LocalGroupName; Status = "FailedToQuery"; Members = $null }
    }
    $memberResults = @()
    foreach ($member in $members) {
        try {
            $MemberName = $member.GetType().InvokeMember("Name","GetProperty",$null,$member,$null)
            $MemberType = $member.GetType().InvokeMember("Class","GetProperty",$null,$member,$null)
            $MemberPath = $member.GetType().InvokeMember("ADSPath","GetProperty",$null,$member,$null)
            $MemberDomain = $null
            if ($MemberPath -match "^Winnt\:\/\/(?<domainName>\S+)\/(?<CompName>\S+)\/") {
                if ($MemberType -eq "User") { $MemberType = "LocalUser" } elseif ($MemberType -eq "Group") { $MemberType = "LocalGroup" }
                $MemberDomain = $matches["CompName"]
            } elseif ($MemberPath -match "^WinNT\:\/\/(?<domainname>\S+)\/") {
                if ($MemberType -eq "User") { $MemberType = "DomainUser" } elseif ($MemberType -eq "Group") { $MemberType = "DomainGroup" }
                $MemberDomain = $matches["domainname"]
            } else {
                $MemberType = "Unknown"
                $MemberDomain = "Unknown"
            }
            $memberResults += "$Computer, $LocalGroupName, SUCCESS, $MemberType, $MemberDomain, $MemberName"
        }
        catch {
            $memberResults += "$Computer, $LocalGroupName, FailedQueryMember"
        }
    }
    Write-Log "Finished scanning $Computer."
    return @{ Computer = $Computer; LocalGroupName = $LocalGroupName; Status = "SUCCESS"; Members = $memberResults }
}

# --- Button Event Handler: Run Audit ---
$btnRunAudit.Add_Click({
    # Save current machine list before scanning
    Save-MachineList
    $global:StopScan = $false
    $selectedMachines = @()
    for ($i=0; $i -lt $clbMachines.Items.Count; $i++) {
        if ($clbMachines.GetItemChecked($i)) { $selectedMachines += $clbMachines.Items[$i].ToString() }
    }
    if ($selectedMachines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No machines selected for scanning.", "Run Audit")
        return
    }
    $LocalGroupName = $txtLocalGroup.Text
    $rescanHours = 2
    [void][Double]::TryParse($txtRescan.Text, [ref]$rescanHours)
    $rescanIntervalSec = [int]($rescanHours * 3600)
    
    # Disable controls during scan
    $clbMachines.Enabled = $false
    $btnRunAudit.Enabled = $false
    $btnGetAD.Enabled = $false
    $btnUp.Enabled = $false
    $btnDown.Enabled = $false
    $btnDelete.Enabled = $false
    $btnSaveList.Enabled = $false
    $btnStopScan.Enabled = $true

    # Initialize CSV files
    Remove-Item $successCSV -ErrorAction SilentlyContinue
    Remove-Item $failureCSV -ErrorAction SilentlyContinue
    "Computer,LocalGroup,Status,Details" | Out-File -FilePath $successCSV -Encoding UTF8
    "Computer,LocalGroup,Status,Details" | Out-File -FilePath $failureCSV -Encoding UTF8

    do {
        $failedMachines = @()
        $machineIndex = 0
        $progressBar.Value = 0
        $progressBar.Maximum = $selectedMachines.Count
        $statusLabel.Text = "Scanning $($selectedMachines.Count) machine(s)..."
        Write-Log "Scanning $($selectedMachines.Count) machine(s)..."
        foreach ($machine in $selectedMachines) {
            if ($global:StopScan) { 
                Write-Log "Scan stopped by user."
                break
            }
            $machineIndex++
            $statusLabel.Text = "Scanning $machine ($machineIndex / $($selectedMachines.Count))..."
            Write-Log "Scanning $machine ($machineIndex / $($selectedMachines.Count))..."
            $scanResult = Scan-Computer -Computer $machine -LocalGroupName $LocalGroupName
            if ($scanResult.Status -eq "SUCCESS") {
                foreach ($line in $scanResult.Members) { 
                    Append-ToCSV -FilePath $successCSV -Line $line 
                }
            } else {
                Append-ToCSV -FilePath $failureCSV -Line "$machine, $LocalGroupName, $($scanResult.Status), "
                $failedMachines += $machine
            }
            $progressBar.Value = $machineIndex
        }
        Write-Log "Completed scanning iteration for $($selectedMachines.Count) machine(s)."
        if ($global:StopScan) {
            Write-Log "Scan stopped by user during iteration."
            break
        }
        if ($failedMachines.Count -gt 0) {
            Write-Log "Rescanning failed machines..."
            # Update computers.config with only the failed machines
            $failedMachines | Out-File -FilePath $computersConfigPath -Encoding UTF8
            # Update the CheckedListBox
            $clbMachines.Items.Clear()
            foreach ($machine in $failedMachines) { [void]$clbMachines.Items.Add($machine) }
            for ($i=0; $i -lt $clbMachines.Items.Count; $i++) { $clbMachines.SetItemChecked($i, $true) }
            # Set selectedMachines for next iteration to the failed machines
            $selectedMachines = $failedMachines
        }
    } while (($failedMachines.Count -gt 0) -and (-not $global:StopScan))
    
    $statusLabel.Text = "Audit complete or stopped."
    Write-Log "Audit complete or stopped."

    # Re-enable controls
    $clbMachines.Enabled = $true
    $btnRunAudit.Enabled = $true
    $btnGetAD.Enabled = $true
    $btnUp.Enabled = $true
    $btnDown.Enabled = $true
    $btnDelete.Enabled = $true
    $btnSaveList.Enabled = $true
    $btnStopScan.Enabled = $true

    # --- Show final results in a new window ---
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Audit Results"
    $resultsForm.Size = New-Object System.Drawing.Size(800,400)
    $resultsForm.StartPosition = "CenterScreen"

    $dataGrid = New-Object System.Windows.Forms.DataGridView
    $dataGrid.Location = New-Object System.Drawing.Point(10,10)
    $dataGrid.Size = New-Object System.Drawing.Size(760,340)
    $dataGrid.AutoSizeColumnsMode = 'Fill'
    $resultsForm.Controls.Add($dataGrid)

    $dt = New-Object System.Data.DataTable
    $dt.Columns.Add("Computer") | Out-Null
    $dt.Columns.Add("LocalGroup") | Out-Null
    $dt.Columns.Add("Status") | Out-Null
    $dt.Columns.Add("Details") | Out-Null

    $lines = Get-Content $successCSV
    if ($lines.Count -gt 1) {
        foreach ($line in $lines[1..($lines.Count - 1)]) {
            $parts = $line.Split(",")
            $row = $dt.NewRow()
            $row["Computer"] = $parts[0].Trim()
            $row["LocalGroup"] = $parts[1].Trim()
            $row["Status"] = "SUCCESS"
            $row["Details"] = if ($parts.Count -ge 4 -and $parts[3]) { $parts[3].Trim() } else { "" }
            $dt.Rows.Add($row) | Out-Null
        }
    }
    $lines = Get-Content $failureCSV
    if ($lines.Count -gt 1) {
        foreach ($line in $lines[1..($lines.Count - 1)]) {
            $parts = $line.Split(",")
            $row = $dt.NewRow()
            $row["Computer"] = $parts[0].Trim()
            $row["LocalGroup"] = $parts[1].Trim()
            $row["Status"] = "FAILED"
            $row["Details"] = if ($parts.Count -ge 4 -and $parts[3]) { $parts[3].Trim() } else { "" }
            $dt.Rows.Add($row) | Out-Null
        }
    }
    $dataGrid.DataSource = $dt

    $resultsForm.ShowDialog() | Out-Null
})

# --- Button Event Handler: Stop Scan ---
$btnStopScan.Add_Click({
    $global:StopScan = $true
    Write-Log "User requested to stop scan."
})

# --- Form Shown Event ---
$form.Add_Shown({ $statusLabel.Text = "Loaded $($clbMachines.Items.Count) machines. Ready to scan." })
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog() | Out-Null

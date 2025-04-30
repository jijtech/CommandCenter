Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Globals ---
$script:scriptsFolder  = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "scripts"
$script:starredFile    = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "starred.txt"
$script:starredScripts = if (Test-Path $script:starredFile) { Get-Content $script:starredFile } else { @() }
$script:Domain         = "proofficepark.dk"
$script:FirstThreeOctets = "192.168.100."
$script:markedComputers = @()

# --- Utility Functions ---
function Save-StarredScripts {
    $script:starredScripts = $script:starredScripts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $script:starredScripts | Set-Content $script:starredFile
}

function Normalize-Path([string]$path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return "" }
    try { [System.IO.Path]::GetFullPath($path).TrimEnd('\').ToLowerInvariant() }
    catch { $path.ToLowerInvariant() }
}

function Get-ScriptGenre([string]$scriptName) {
    switch -Regex ($scriptName) {
        "GetAD"      { "Active Directory"; break }
        "Compliance" { "Compliance"; break }
        default      { "Local" }
    }
}

function Get-ScriptSubGenre([string]$scriptName) {
    switch -Regex ($scriptName) {
        "(?i)enable"   { "Enable"; break }
        "(?i)disable"  { "Disable"; break }
        "(?i)audit"    { "Audit"; break }
        "(?i)security" { "Security"; break }
        "(?i)logs?"    { "Logs"; break }
        "(?i)setup"    { "Setup"; break }
        default        { "Other" }
    }
}

function Get-SelectedSubGenre {
    if ($otherSubGenreRadio.Checked) { "Other" }
    elseif ($enableSubGenreRadio.Checked) { "Enable" }
    elseif ($disableSubGenreRadio.Checked) { "Disable" }
    elseif ($auditSubGenreRadio -and $auditSubGenreRadio.Checked) { "Audit" }
    elseif ($securitySubGenreRadio -and $securitySubGenreRadio.Checked) { "Security" }
    elseif ($logsSubGenreRadio -and $logsSubGenreRadio.Checked) { "Logs" }
    elseif ($setupSubGenreRadio -and $setupSubGenreRadio.Checked) { "Setup" }
    else { "Other" }
}

# --- UI Construction ---
# Main Form
$mainForm = New-Object System.Windows.Forms.Form -Property @{
    Text = "CommandCenter"
    Size = [System.Drawing.Size]::new(1200, 650)
    StartPosition = "CenterScreen"
}

# Main Layout
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    Dock = 'Fill'; RowCount = 1; ColumnCount = 2
}
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

# --- Command Panel (Left) ---
$commandPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = 'Fill' }

# Change: Add an extra row for the second sub-genre row
$tableLayout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
    Dock = 'Fill'; RowCount = 6; ColumnCount = 1
}
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 8)))   # For separator line
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  # New row for second sub-genre row
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# Buttons Panel
$buttonsPanel = New-Object System.Windows.Forms.Panel -Property @{
    Dock = 'Fill'; Padding = [System.Windows.Forms.Padding]::new(10)
}
$selectFolderButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Select Scripts Folder"; Location = [System.Drawing.Point]::new(10, 10); Size = [System.Drawing.Size]::new(150, 30)
}
$refreshButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Refresh"; Location = [System.Drawing.Point]::new(170, 10); Size = [System.Drawing.Size]::new(100, 30)
}
$runScriptButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Run Selected Script"; Location = [System.Drawing.Point]::new(10, 45); Size = [System.Drawing.Size]::new(150, 30)
}
$adCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = "Active Directory"; Location = [System.Drawing.Point]::new(170, 50); AutoSize = $true
}
$localCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = "Local"; Location = [System.Drawing.Point]::new(300, 50); AutoSize = $true
}
$complianceCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = "Compliance"; Location = [System.Drawing.Point]::new(400, 50); AutoSize = $true
}
$folderLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Current Folder: $script:scriptsFolder"; Location = [System.Drawing.Point]::new(330, 15); AutoSize = $true; Height = 20
}
$buttonsPanel.Controls.AddRange(@(
    $selectFolderButton, $refreshButton, $runScriptButton,
    $adCheckBox, $localCheckBox, $complianceCheckBox, $folderLabel
))

# ---- Separator Panel
$separatorPanel = New-Object System.Windows.Forms.Panel -Property @{
    Dock = 'Fill'; Height = 8; Margin = [System.Windows.Forms.Padding]::new(0,0,0,0)
    BackColor = [System.Drawing.Color]::FromArgb(180,180,180)
}
$separatorLine = New-Object System.Windows.Forms.Label -Property @{
    BorderStyle = 'Fixed3D'
    AutoSize = $false
    Height = 1   # Make the separator line thinner (was 2)
    Width = 900
    Dock = 'Top'
    BackColor = [System.Drawing.Color]::FromArgb(180,180,180)
    Margin = [System.Windows.Forms.Padding]::new(0,3,0,3)
}
$separatorPanel.Controls.Add($separatorLine)

# Sub-genre Panel Row 1
$subGenrePanel1 = New-Object System.Windows.Forms.Panel -Property @{
    Dock = 'Fill'; Height = 35; Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
}
$subGenreLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Sub-genre:"; Location = [System.Drawing.Point]::new(10, 7); AutoSize = $true; Height = 20
}
# Row 1: Other, Enable, Disable, Audit
$otherSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Other"; Location = [System.Drawing.Point]::new(90, 5); AutoSize = $true; Checked = $true; Name = "OtherSubGenreRadio"
}
$enableSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Enable"; Location = [System.Drawing.Point]::new(170, 5); AutoSize = $true; Checked = $false; Name = "EnableSubGenreRadio"
}
$disableSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Disable"; Location = [System.Drawing.Point]::new(260, 5); AutoSize = $true; Checked = $false; Name = "DisableSubGenreRadio"
}
$auditSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Audit"; Location = [System.Drawing.Point]::new(350, 5); AutoSize = $true; Checked = $false; Name = "AuditSubGenreRadio"
}
$subGenrePanel1.Controls.AddRange(@(
    $subGenreLabel, $otherSubGenreRadio, $enableSubGenreRadio, $disableSubGenreRadio, $auditSubGenreRadio
))

# Sub-genre Panel Row 2
$subGenrePanel2 = New-Object System.Windows.Forms.Panel -Property @{
    Dock = 'Fill'; Height = 35; Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
}
# Row 2: Security, Logs, Setup
$securitySubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Security"; Location = [System.Drawing.Point]::new(90, 5); AutoSize = $true; Checked = $false; Name = "SecuritySubGenreRadio"
}
$logsSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Logs"; Location = [System.Drawing.Point]::new(200, 5); AutoSize = $true; Checked = $false; Name = "LogsSubGenreRadio"
}
$setupSubGenreRadio = New-Object System.Windows.Forms.RadioButton -Property @{
    Text = "Setup"; Location = [System.Drawing.Point]::new(300, 5); AutoSize = $true; Checked = $false; Name = "SetupSubGenreRadio"
}
$subGenrePanel2.Controls.AddRange(@(
    $securitySubGenreRadio, $logsSubGenreRadio, $setupSubGenreRadio
))

# Search Panel
$searchPanel = New-Object System.Windows.Forms.Panel -Property @{
    Dock = 'Fill'; Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
}
$searchLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Search:"; Location = [System.Drawing.Point]::new(10, 5); AutoSize = $true; Height = 20
}
$searchTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = [System.Drawing.Point]::new(60, 3); Size = [System.Drawing.Size]::new(700, 20)
}
$searchPanel.Controls.AddRange(@($searchLabel, $searchTextBox))

# Scripts ListView
$scriptsListView = New-Object System.Windows.Forms.ListView -Property @{
    View = [System.Windows.Forms.View]::Details
    Dock = 'Fill'
    FullRowSelect = $true
    GridLines = $true
    MultiSelect = $false
    HideSelection = $false
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::Black
    Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    CheckBoxes = $true
}
$scriptsListView.Columns.Clear()
$scriptsListView.Columns.AddRange(@(
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Favorite"; Width = 60 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Script Name"; Width = 250 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Description"; Width = 300 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Last Modified"; Width = 150 })
))

# Compose Command Panel
$tableLayout.Controls.Add($buttonsPanel, 0, 0)
$tableLayout.Controls.Add($separatorPanel, 0, 1)
$tableLayout.Controls.Add($subGenrePanel1, 0, 2)
$tableLayout.Controls.Add($subGenrePanel2, 0, 3)
$tableLayout.Controls.Add($searchPanel, 0, 4)
$tableLayout.Controls.Add($scriptsListView, 0, 5)
$commandPanel.Controls.Add($tableLayout)

# --- Script List Logic ---
function Select-ScriptsList {
    $searchText = $searchTextBox.Text.ToLower()
    $selectedSubGenre = Get-SelectedSubGenre
    $scriptsListView.BeginUpdate()
    $scriptsListView.Items.Clear()
    if ($script:allScripts) {
        $normalizedStarred = @{}
        foreach ($starred in $script:starredScripts) {
            $norm = Normalize-Path $starred
            if ($norm) { $normalizedStarred[$norm] = $true }
        }
        foreach ($script in $script:allScripts) {
            $genre = $script.Genre
            $subGenre = $script.SubGenre
            $show = $false
            if ($adCheckBox.Checked -and $genre -eq "Active Directory") { $show = $true }
            if ($localCheckBox.Checked -and $genre -eq "Local") { $show = $true }
            if ($complianceCheckBox.Checked -and $genre -eq "Compliance") { $show = $true }
            if (-not $adCheckBox.Checked -and -not $localCheckBox.Checked -and -not $complianceCheckBox.Checked) { $show = $true }
            if ($show -and $selectedSubGenre -and $selectedSubGenre -ne "" -and $subGenre -ne $selectedSubGenre) { $show = $false }
            if ($show -and (
                $script.Name.ToLower().Contains($searchText) -or
                $script.Description.ToLower().Contains($searchText)
            )) {
                $normScript = Normalize-Path $script.FullPath
                $isStarred = $normalizedStarred.ContainsKey($normScript)
                $item = New-Object System.Windows.Forms.ListViewItem("")
                $item.Checked = $isStarred
                $item.SubItems.Add($script.Name) | Out-Null
                $descSubItem = $item.SubItems.Add($script.Description)
                $descSubItem.ForeColor = [System.Drawing.Color]::DarkBlue
                $dateSubItem = $item.SubItems.Add($script.LastModified)
                $dateSubItem.ForeColor = [System.Drawing.Color]::DarkGreen
                $item.Tag = $script.FullPath
                $scriptsListView.Items.Add($item) | Out-Null
            }
        }
    }
    $scriptsListView.EndUpdate()
}

function Update-ScriptsList {
    $script:allScripts = @()
    $scriptsListView.BeginUpdate()
    $scriptsListView.Items.Clear()
    $folderLabel.Text = "Current Folder: $script:scriptsFolder"
    try {
        $scripts = Get-ChildItem -Path $script:scriptsFolder -Filter "*.ps1" -Recurse -ErrorAction Stop
        $normalizedStarred = @{}
        foreach ($starred in $script:starredScripts) {
            $norm = Normalize-Path $starred
            if ($norm) { $normalizedStarred[$norm] = $true }
        }
        foreach ($script in $scripts) {
            try {
                $description = ""
                $content = Get-Content -Path $script.FullName -TotalCount 10 -ErrorAction Stop
                foreach ($line in $content) {
                    if ($line -match "^#\s*Description:\s*(.+)$") {
                        $description = $matches[1].Trim()
                        break
                    }
                }
                $scriptInfo = [pscustomobject]@{
                    Name        = $script.Name
                    Description = $description
                    LastModified= $script.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    FullPath    = $script.FullName
                    Genre       = Get-ScriptGenre $script.Name
                    SubGenre    = Get-ScriptSubGenre $script.Name
                }
                $script:allScripts += $scriptInfo
                $normScript = Normalize-Path $script.FullName
                $isStarred = $normalizedStarred.ContainsKey($normScript)
                $item = New-Object System.Windows.Forms.ListViewItem("")
                $item.Checked = $isStarred
                $item.SubItems.Add($script.Name) | Out-Null
                $descSubItem = $item.SubItems.Add($description)
                $descSubItem.ForeColor = [System.Drawing.Color]::DarkBlue
                $dateSubItem = $item.SubItems.Add($script.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
                $dateSubItem.ForeColor = [System.Drawing.Color]::DarkGreen
                $item.Tag = $script.FullName
                $scriptsListView.Items.Add($item) | Out-Null
            }
            catch {
                $scriptInfo = [pscustomobject]@{
                    Name        = $script.Name
                    Description = "Error reading script"
                    LastModified= $script.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    FullPath    = $script.FullName
                    Genre       = Get-ScriptGenre $script.Name
                    SubGenre    = Get-ScriptSubGenre $script.Name
                }
                $script:allScripts += $scriptInfo
                $item = New-Object System.Windows.Forms.ListViewItem("")
                $item.Checked = $false
                $item.SubItems.Add($script.Name) | Out-Null
                $descSubItem = $item.SubItems.Add("Error reading script")
                $descSubItem.ForeColor = [System.Drawing.Color]::Red
                $dateSubItem = $item.SubItems.Add($script.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
                $dateSubItem.ForeColor = [System.Drawing.Color]::DarkGreen
                $item.Tag = $script.FullName
                $scriptsListView.Items.Add($item) | Out-Null
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error accessing scripts directory: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    $scriptsListView.EndUpdate()
    $scriptsListView.Refresh()
}

# --- Event Handlers ---
$searchTextBox.Add_TextChanged({ Select-ScriptsList })
$adCheckBox.Add_CheckedChanged({ Select-ScriptsList })
$localCheckBox.Add_CheckedChanged({ Select-ScriptsList })
$complianceCheckBox.Add_CheckedChanged({ Select-ScriptsList })
$enableSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$disableSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$otherSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$auditSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$securitySubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$logsSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })
$setupSubGenreRadio.Add_CheckedChanged({ Select-ScriptsList })

$selectFolderButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Scripts Folder"
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $script:scriptsFolder = $folderBrowser.SelectedPath
        Update-ScriptsList
    }
})

$refreshButton.Add_Click({ Update-ScriptsList })

$runScriptButton.Add_Click({
    if ($scriptsListView.SelectedItems.Count -gt 0) {
        $scriptPath = $scriptsListView.SelectedItems[0].Tag
        try {
            Start-Process powershell.exe -ArgumentList "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error executing script: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$scriptsListView.Add_ItemCheck({
    param($sender, $e)
    $item = $scriptsListView.Items[$e.Index]
    $scriptPath = $item.Tag
    $normalizedScriptPath = Normalize-Path $scriptPath
    $script:starredScripts = $script:starredScripts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $normalizedStarred = @{}
    foreach ($starred in $script:starredScripts) {
        $norm = Normalize-Path $starred
        if ($norm) { $normalizedStarred[$norm] = $starred }
    }
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        if (-not $normalizedStarred.ContainsKey($normalizedScriptPath)) {
            $script:starredScripts += $scriptPath
            Save-StarredScripts
        }
    } else {
        if ($normalizedStarred.ContainsKey($normalizedScriptPath)) {
            $script:starredScripts = $script:starredScripts | Where-Object { Normalize-Path $_ -ne $normalizedScriptPath }
            Save-StarredScripts
        }
    }
})

Update-ScriptsList

# --- Network Scanner Panel (Right) ---
$scannerPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = 'Fill' }

# Top Controls
$scannerTopPanel = New-Object System.Windows.Forms.Panel -Property @{ Dock = 'Top'; Height = 40 }
$ipLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "First 3 Octets:"; Location = [System.Drawing.Point]::new(10, 10); AutoSize = $true
}
$ipTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = [System.Drawing.Point]::new(100, 7); Size = [System.Drawing.Size]::new(110, 22); Text = $script:FirstThreeOctets
}
$domainLabel = New-Object System.Windows.Forms.Label -Property @{
    Text = "Domain:"; Location = [System.Drawing.Point]::new(230, 10); AutoSize = $true
}
$domainTextBox = New-Object System.Windows.Forms.TextBox -Property @{
    Location = [System.Drawing.Point]::new(290, 7); Size = [System.Drawing.Size]::new(180, 22); Text = $script:Domain
}
$rescanButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Rescan"; Location = [System.Drawing.Point]::new(490, 5); Size = [System.Drawing.Size]::new(90, 28)
}
$scannerTopPanel.Controls.AddRange(@($ipLabel, $ipTextBox, $domainLabel, $domainTextBox, $rescanButton))

# Computers ListView
$computersListView = New-Object System.Windows.Forms.ListView -Property @{
    View = [System.Windows.Forms.View]::Details
    Dock = 'Top'
    Height = 450
    FullRowSelect = $true
    GridLines = $true
    MultiSelect = $false
    HideSelection = $false
    BackColor = [System.Drawing.Color]::White
    ForeColor = [System.Drawing.Color]::Black
    Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    CheckBoxes = $true
}
$computersListView.Columns.Clear()
$computersListView.Columns.AddRange(@(
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Mark"; Width = 50 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Type"; Width = 120 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "IP Address"; Width = 120 }),
    (New-Object System.Windows.Forms.ColumnHeader -Property @{ Text = "Host Name"; Width = 250 })
))

# Progress Bar and Percent Label
$progressBar = New-Object System.Windows.Forms.ProgressBar -Property @{
    Dock = 'Top'; Height = 30; Minimum = 0; Maximum = 254; Value = 0
}
$percentLabel = New-Object System.Windows.Forms.Label -Property @{
    Dock = 'Top'; Height = 25; TextAlign = 'MiddleCenter'; Text = "0%"
}
$closeButton = New-Object System.Windows.Forms.Button -Property @{
    Text = "Close"; Dock = 'Bottom'; Enabled = $false
}
$closeButton.Add_Click({ $mainForm.Close() })

# Compose Scanner Panel
$scannerPanel.Controls.Add($closeButton)
$scannerPanel.Controls.Add($percentLabel)
$scannerPanel.Controls.Add($progressBar)
$scannerPanel.Controls.Add($computersListView)
$scannerPanel.Controls.Add($scannerTopPanel)

# --- Scan Logic ---
$scanState = @{
    Current = 1
    Results = @()
    ScanCompleted = $false
}

function Scan-NextIP {
    if ($scanState.ScanCompleted) { return }
    if ($scanState.Current -gt 254) {
        $percentLabel.Text = "Scan complete!"
        $progressBar.Value = 254
        $closeButton.Enabled = $true
        $scanState.ScanCompleted = $true
        $timer.Stop()
        return
    }
    $i = $scanState.Current
    $IP = "$script:FirstThreeOctets$i"
    $entry = $null
    $hostType = ""
    $hostName = ""
    try {
        $ping = [System.Net.NetworkInformation.Ping]::new()
        $pingResults = $ping.Send($IP, 300)
    } catch { $pingResults = $null }
    if ($pingResults -and $pingResults.Status -eq "Success") {
        try {
            $DnsName = [System.Net.Dns]::GetHostEntry($IP)
        } catch { $DnsName = $null }
        if ($DnsName -and $DnsName.HostName) {
            $hostType = if ($DnsName.HostName.EndsWith(".$script:Domain")) { "DOMAIN_HOST_FOUND" } else { "NON_DOMAIN_HOST_FOUND" }
            $hostName = $DnsName.HostName
            $entry = "[$hostType] $IP -> $hostName"
        }
    }
    if ($entry) {
        $scanState.Results += $entry
        $item = New-Object System.Windows.Forms.ListViewItem("")
        $item.Checked = $false
        $item.SubItems.Add($hostType) | Out-Null
        $item.SubItems.Add($IP) | Out-Null
        $item.SubItems.Add($hostName) | Out-Null
        $item.Tag = @{ IP = $IP; HostName = $hostName; Type = $hostType }
        $computersListView.Items.Add($item) | Out-Null
        $computersListView.TopItem = $item
    }
    $progressBar.Value = $i
    $percentLabel.Text = ("{0}%" -f ([math]::Round(($i/254)*100)))
    $scanState.Current++
}

$computersListView.Add_ItemCheck({
    param($sender, $e)
    $item = $computersListView.Items[$e.Index]
    $tag = $item.Tag
    if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
        if (-not $script:markedComputers | Where-Object { $_.IP -eq $tag.IP -and $_.HostName -eq $tag.HostName }) {
            $script:markedComputers += $tag
        }
    } else {
        $script:markedComputers = $script:markedComputers | Where-Object { $_.IP -ne $tag.IP -or $_.HostName -ne $tag.HostName }
    }
})

# --- Timer for async scan ---
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30
$timer.Add_Tick({ Scan-NextIP })

function Start-Scan {
    $script:FirstThreeOctets = $ipTextBox.Text.Trim()
    if ($script:FirstThreeOctets -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}\.$') {
        [System.Windows.Forms.MessageBox]::Show("First 3 octets must be in the format 'X.X.X.' (ending with a dot)", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $script:Domain = $domainTextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($script:Domain)) {
        [System.Windows.Forms.MessageBox]::Show("Domain cannot be empty.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $scanState.Current = 1
    $scanState.Results = @()
    $scanState.ScanCompleted = $false
    $computersListView.Items.Clear()
    $progressBar.Value = 0
    $percentLabel.Text = "0%"
    $closeButton.Enabled = $false
    $timer.Start()
}

$mainForm.Add_Shown({ Start-Scan })
$rescanButton.Add_Click({ Start-Scan })

# --- Compose Main Layout ---
$mainLayout.Controls.Add($commandPanel, 0, 0)
$mainLayout.Controls.Add($scannerPanel, 1, 0)
$mainForm.Controls.Add($mainLayout)

# --- Ensure scripts folder exists ---
if (!(Test-Path $script:scriptsFolder)) {
    New-Item -ItemType Directory -Path $script:scriptsFolder -Force | Out-Null
    Write-Host "Created scripts folder: $script:scriptsFolder"
}

# --- Show Main Form ---
[void]$mainForm.ShowDialog()
Add-Type -AssemblyName System.Windows.Forms

$script:scriptsFolder  = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "scripts"
$script:starredFile    = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "starred.txt"
$script:starredScripts = if (Test-Path $script:starredFile) { Get-Content $script:starredFile } else { @() }

function Save-StarredScripts {
    $script:starredScripts = $script:starredScripts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $script:starredScripts | Set-Content $script:starredFile
}

function Show-ScriptMenuGUI {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text = "CommandCenter"
        Size = [System.Drawing.Size]::new(800, 600)
    }

    $tableLayout = New-Object System.Windows.Forms.TableLayoutPanel -Property @{
        Dock = 'Fill'
        RowCount = 3
        ColumnCount = 1
    }
    $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
    $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    $tableLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $tableLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $buttonsPanel = New-Object System.Windows.Forms.Panel -Property @{
        Dock = 'Fill'
        Padding = [System.Windows.Forms.Padding]::new(10)
    }

    $selectFolderButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Select Scripts Folder"
        Location = [System.Drawing.Point]::new(10, 10)
        Size = [System.Drawing.Size]::new(150, 30)
    }
    $refreshButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Refresh"
        Location = [System.Drawing.Point]::new(170, 10)
        Size = [System.Drawing.Size]::new(100, 30)
    }
    $runScriptButton = New-Object System.Windows.Forms.Button -Property @{
        Text = "Run Selected Script"
        Location = [System.Drawing.Point]::new(10, 45)
        Size = [System.Drawing.Size]::new(150, 30)
    }

    $adCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Active Directory"
        Location = [System.Drawing.Point]::new(170, 50)
        AutoSize = $true
    }
    $localCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Local"
        Location = [System.Drawing.Point]::new(300, 50)
        AutoSize = $true
    }
    $complianceCheckBox = New-Object System.Windows.Forms.CheckBox -Property @{
        Text = "Compliance"
        Location = [System.Drawing.Point]::new(400, 50)
        AutoSize = $true
    }

    $folderLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Current Folder: $script:scriptsFolder"
        Location = [System.Drawing.Point]::new(330, 15)
        AutoSize = $true
        Height = 20
    }

    $searchPanel = New-Object System.Windows.Forms.Panel -Property @{
        Dock = 'Fill'
        Padding = [System.Windows.Forms.Padding]::new(10, 0, 10, 0)
    }
    $searchLabel = New-Object System.Windows.Forms.Label -Property @{
        Text = "Search:"
        Location = [System.Drawing.Point]::new(10, 5)
        AutoSize = $true
        Height = 20
    }
    $searchTextBox = New-Object System.Windows.Forms.TextBox -Property @{
        Location = [System.Drawing.Point]::new(60, 3)
        Size = [System.Drawing.Size]::new(700, 20)
    }

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

    $buttonsPanel.Controls.AddRange(@(
        $selectFolderButton, $refreshButton, $runScriptButton,
        $adCheckBox, $localCheckBox, $complianceCheckBox, $folderLabel
    ))
    $searchPanel.Controls.AddRange(@($searchLabel, $searchTextBox))
    $tableLayout.Controls.Add($buttonsPanel, 0, 0)
    $tableLayout.Controls.Add($searchPanel, 0, 1)
    $tableLayout.Controls.Add($scriptsListView, 0, 2)
    $form.Controls.Add($tableLayout)

    function Normalize-Path([string]$path) {
        if ([string]::IsNullOrWhiteSpace($path)) { return "" }
        try { [System.IO.Path]::GetFullPath($path).TrimEnd('\').ToLowerInvariant() }
        catch { $path.ToLowerInvariant() }
    }

    function Get-ScriptGenre([string]$scriptName) {
        if ($scriptName -match "GetAD") { "Active Directory" }
        elseif ($scriptName -match "Compliance") { "Compliance" }
        else { "Local" }
    }

    function Select-ScriptsList {
        $searchText = $searchTextBox.Text.ToLower()
        $scriptsListView.BeginUpdate()
        $scriptsListView.Items.Clear()
        if ($script:allScripts) {
            $normalizedStarred = @{}
            foreach ($starred in $script:starredScripts) {
                $norm = Normalize-Path $starred
                if ($norm) { $normalizedStarred[$norm] = $true }
            }
            foreach ($script in $script:allScripts) {
                $genre = Get-ScriptGenre $script.Name
                $show = $false
                if ($adCheckBox.Checked -and $genre -eq "Active Directory") { $show = $true }
                if ($localCheckBox.Checked -and $genre -eq "Local") { $show = $true }
                if ($complianceCheckBox.Checked -and $genre -eq "Compliance") { $show = $true }
                if (-not $adCheckBox.Checked -and -not $localCheckBox.Checked -and -not $complianceCheckBox.Checked) { $show = $true }
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

    $searchTextBox.Add_TextChanged({ Select-ScriptsList })
    $adCheckBox.Add_CheckedChanged({ Select-ScriptsList })
    $localCheckBox.Add_CheckedChanged({ Select-ScriptsList })
    $complianceCheckBox.Add_CheckedChanged({ Select-ScriptsList })

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

    $form.Add_Shown({
        $form.Activate()
        Update-ScriptsList
    })

    [void]$form.ShowDialog()
}

if (!(Test-Path $script:scriptsFolder)) {
    New-Item -ItemType Directory -Path $script:scriptsFolder -Force | Out-Null
    Write-Host "Created scripts folder: $script:scriptsFolder"
}

Show-ScriptMenuGUI
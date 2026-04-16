Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- MAIN FORM ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Profit-Track Data Formatter ~ By James Connolly - 2026"
$form.Size = New-Object System.Drawing.Size(1100,700)
$form.MinimumSize = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

# ---------------- TOP PANEL ----------------
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Location = New-Object System.Drawing.Point(20,15)
$topPanel.Size = New-Object System.Drawing.Size(1040,95)
$topPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($topPanel)

# ---------------- FILE INPUT ROW ----------------
$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point(0,0)
$inputBox.Size = New-Object System.Drawing.Size(845,25)
$inputBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$topPanel.Controls.Add($inputBox)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Size = New-Object System.Drawing.Size(80,27)
$btnBrowse.Location = New-Object System.Drawing.Point(780,0)
$btnBrowse.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$topPanel.Controls.Add($btnBrowse)

$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Process"
$btnProcess.Size = New-Object System.Drawing.Size(80,27)
$btnProcess.Location = New-Object System.Drawing.Point(870,0)
$btnProcess.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnProcess.Visible = $false
$topPanel.Controls.Add($btnProcess)

# ---------------- SEARCH ROW ----------------
$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(0,40)
$topPanel.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(55,36)
$txtSearch.Size = New-Object System.Drawing.Size(360,25)
$txtSearch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$topPanel.Controls.Add($txtSearch)

$btnFieldPicker = New-Object System.Windows.Forms.Button
$btnFieldPicker.Text = "Search Fields ▼"
$btnFieldPicker.Size = New-Object System.Drawing.Size(140,27)
$btnFieldPicker.Location = New-Object System.Drawing.Point(430,35)
$btnFieldPicker.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$topPanel.Controls.Add($btnFieldPicker)

$btnClearFilter = New-Object System.Windows.Forms.Button
$btnClearFilter.Text = "Clear Filter"
$btnClearFilter.Size = New-Object System.Drawing.Size(95,27)
$btnClearFilter.Location = New-Object System.Drawing.Point(580,35)
$btnClearFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$topPanel.Controls.Add($btnClearFilter)

$chkExportFiltered = New-Object System.Windows.Forms.CheckBox
$chkExportFiltered.Text = "Export filtered rows only"
$chkExportFiltered.AutoSize = $true
$chkExportFiltered.Location = New-Object System.Drawing.Point(690,39)
$chkExportFiltered.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$chkExportFiltered.Enabled = $false
$topPanel.Controls.Add($chkExportFiltered)

# ---------------- FILTER POPUP PANEL ----------------
$filterPanel = New-Object System.Windows.Forms.Panel
$filterPanel.Size = New-Object System.Drawing.Size(300,260)
$filterPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$filterPanel.Visible = $false
$filterPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($filterPanel)
$filterPanel.BringToFront()

$lblFilterPanel = New-Object System.Windows.Forms.Label
$lblFilterPanel.Text = "Choose one or more fields to search:"
$lblFilterPanel.AutoSize = $true
$lblFilterPanel.Location = New-Object System.Drawing.Point(8,8)
$filterPanel.Controls.Add($lblFilterPanel)

$clbFilterFields = New-Object System.Windows.Forms.CheckedListBox
$clbFilterFields.Location = New-Object System.Drawing.Point(8,30)
$clbFilterFields.Size = New-Object System.Drawing.Size(282,185)
$clbFilterFields.CheckOnClick = $true
$clbFilterFields.IntegralHeight = $false
$clbFilterFields.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$filterPanel.Controls.Add($clbFilterFields)

$btnSelectAllFields = New-Object System.Windows.Forms.Button
$btnSelectAllFields.Text = "All"
$btnSelectAllFields.Size = New-Object System.Drawing.Size(55,25)
$btnSelectAllFields.Location = New-Object System.Drawing.Point(8,225)
$btnSelectAllFields.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$filterPanel.Controls.Add($btnSelectAllFields)

$btnSelectNoneFields = New-Object System.Windows.Forms.Button
$btnSelectNoneFields.Text = "None"
$btnSelectNoneFields.Size = New-Object System.Drawing.Size(55,25)
$btnSelectNoneFields.Location = New-Object System.Drawing.Point(70,225)
$btnSelectNoneFields.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$filterPanel.Controls.Add($btnSelectNoneFields)

$btnCloseFilterPanel = New-Object System.Windows.Forms.Button
$btnCloseFilterPanel.Text = "Close"
$btnCloseFilterPanel.Size = New-Object System.Drawing.Size(60,25)
$btnCloseFilterPanel.Location = New-Object System.Drawing.Point(230,225)
$btnCloseFilterPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$filterPanel.Controls.Add($btnCloseFilterPanel)

# ---------------- GRID ----------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20,120)
$grid.Size = New-Object System.Drawing.Size(1040,440)
$grid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToOrderColumns = $true
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
$grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect = $false
$form.Controls.Add($grid)

# ---------------- PROGRESS BAR ----------------
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20,575)
$progress.Size = New-Object System.Drawing.Size(1040,20)
$progress.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progress)

# ---------------- LABEL ----------------
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(20,600)
$lblProgress.Size = New-Object System.Drawing.Size(760,20)
$lblProgress.Text = "Ready"
$lblProgress.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lblProgress)

# ---------------- EXPORT BUTTONS ----------------
$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Text = "Export TXT"
$btnExportTxt.Size = New-Object System.Drawing.Size(95,28)
$btnExportTxt.Location = New-Object System.Drawing.Point(20,630)
$btnExportTxt.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnExportTxt)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Size = New-Object System.Drawing.Size(95,28)
$btnExportCsv.Location = New-Object System.Drawing.Point(125,630)
$btnExportCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnExportCsv)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Export"
$btnCancel.Size = New-Object System.Drawing.Size(110,28)
$btnCancel.Location = New-Object System.Drawing.Point(235,630)
$btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnCancel)

# ---------------- DRAG & DROP ----------------
$form.AllowDrop = $true
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = [Windows.Forms.DragDropEffects]::Copy
    }
})

$form.Add_DragDrop({
    $file = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)[0]
    $inputBox.Text = $file
    Start-ParseIfValid $file
})

# ---------------- STORAGE ----------------
$script:dataTable = New-Object System.Data.DataTable
$script:dataView = $null
$script:cancelExport = $false
$script:searchDelayTimer = New-Object System.Windows.Forms.Timer
$script:searchDelayTimer.Interval = 300
$script:highlightSearchText = ""
$script:highlightColumns = @()
$script:suspendFilterEvents = $false

# ---------------- HELPERS ----------------
function Update-TopPanelLayout {
    $padding = 0
    $gap = 10

    $browseWidth = $btnBrowse.Width
    $rightEdgeRow1 = $topPanel.ClientSize.Width - $padding
    $btnBrowse.Location = New-Object System.Drawing.Point(($rightEdgeRow1 - $browseWidth), 0)

    $inputWidth = [Math]::Max(250, $btnBrowse.Left - $gap - $padding)
    $inputBox.Location = New-Object System.Drawing.Point($padding, 0)
    $inputBox.Size = New-Object System.Drawing.Size($inputWidth, 25)

    $labelX = 0
    $searchTop = 36
    $lblSearch.Location = New-Object System.Drawing.Point($labelX, 40)

    $searchX = 55
    $fieldPickerWidth = $btnFieldPicker.Width
    $clearWidth = $btnClearFilter.Width
    $checkboxWidth = $chkExportFiltered.PreferredSize.Width

    $checkboxX = $topPanel.ClientSize.Width - $checkboxWidth
    $clearX = $checkboxX - $gap - $clearWidth
    $fieldPickerX = $clearX - $gap - $fieldPickerWidth
    $searchWidth = [Math]::Max(160, $fieldPickerX - $gap - $searchX)

    $txtSearch.Location = New-Object System.Drawing.Point($searchX, $searchTop)
    $txtSearch.Size = New-Object System.Drawing.Size($searchWidth, 25)

    $btnFieldPicker.Location = New-Object System.Drawing.Point($fieldPickerX, 35)
    $btnClearFilter.Location = New-Object System.Drawing.Point($clearX, 35)
    $chkExportFiltered.Location = New-Object System.Drawing.Point($checkboxX, 39)
}
Update-TopPanelLayout

function Set-Status {
    param(
        [string]$Text,
        [int]$Value = -1,
        [int]$Maximum = -1
    )

    if ($Maximum -ge 0) {
        $progress.Maximum = [Math]::Max(1, $Maximum)
    }

    if ($Value -ge 0) {
        $progress.Value = [Math]::Min($progress.Maximum, [Math]::Max(0, $Value))
    }

    $lblProgress.Text = $Text
}

function Get-ETA {
    param($startTime, $current, $total)

    if ($current -le 0 -or $total -le 0) { return "Calculating..." }

    $elapsed = (Get-Date) - $startTime
    if ($elapsed.TotalSeconds -le 0) { return "Calculating..." }

    $rate = $current / $elapsed.TotalSeconds
    if ($rate -le 0) { return "Calculating..." }

    $remaining = ($total - $current) / $rate
    if ($remaining -lt 0) { $remaining = 0 }

    return ([TimeSpan]::FromSeconds($remaining)).ToString("hh\:mm\:ss")
}

function Update-ProgressThrottled {
    param(
        [datetime]$StartTime,
        [int]$Current,
        [int]$Total,
        [string]$Prefix = "Processing",
        [int]$Step = 200,
        [switch]$Force
    )

    if (-not $Force -and $Current % $Step -ne 0 -and $Current -ne $Total) {
        return
    }

    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    $eta = Get-ETA $StartTime $Current $Total
    Set-Status "$Prefix... $Current / $Total ($percent%) | ETA: $eta" $Current $Total
    [System.Windows.Forms.Application]::DoEvents()
}

function Escape-CsvValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) { return "" }

    $text = [string]$Value
    if ($text.Contains('"')) {
        $text = $text.Replace('"', '""')
    }

    if ($text.Contains(',') -or $text.Contains('"') -or $text.Contains("`r") -or $text.Contains("`n")) {
        return '"' + $text + '"'
    }

    return $text
}

function Normalize-Row {
    param(
        [string[]]$Values,
        [int]$ExpectedCount
    )

    if ($null -eq $Values) {
        $Values = @()
    }

    if ($Values.Count -lt $ExpectedCount) {
        $result = New-Object string[] $ExpectedCount
        for ($i = 0; $i -lt $ExpectedCount; $i++) {
            if ($i -lt $Values.Count) {
                $result[$i] = $Values[$i].Trim()
            }
            else {
                $result[$i] = ""
            }
        }
        return $result
    }

    if ($Values.Count -gt $ExpectedCount) {
        $result = New-Object string[] $ExpectedCount
        for ($i = 0; $i -lt $ExpectedCount; $i++) {
            $result[$i] = $Values[$i].Trim()
        }
        return $result
    }

    for ($i = 0; $i -lt $Values.Count; $i++) {
        $Values[$i] = $Values[$i].Trim()
    }

    return $Values
}

function Start-ParseIfValid {
    param([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        return
    }

    if (Test-Path -LiteralPath $FilePath) {
        Parse-File $FilePath
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please choose a valid file first.", "Missing file", "OK", "Warning")
    }
}


function Escape-RowFilterValue {
    param([string]$Value)

    if ($null -eq $Value) { return "" }

    $escaped = $Value.Replace("'", "''")
    $escaped = $escaped.Replace("[", "[[]")
    $escaped = $escaped.Replace("]", "[]]")
    $escaped = $escaped.Replace("%", "[%]")
    $escaped = $escaped.Replace("*", "[*]")
    return $escaped
}

function Update-FilterButtonText {
    if ($clbFilterFields.CheckedItems.Count -eq 0) {
        $btnFieldPicker.Text = "Search Fields ▼"
        return
    }

    if ($clbFilterFields.CheckedItems.Count -eq $clbFilterFields.Items.Count) {
        $btnFieldPicker.Text = "Search Fields (All) ▼"
        return
    }

    $btnFieldPicker.Text = "Search Fields ($($clbFilterFields.CheckedItems.Count)) ▼"
}

function Initialize-FilterColumns {
    $script:suspendFilterEvents = $true
    try {
        $clbFilterFields.Items.Clear()

        $columnNames = New-Object System.Collections.Generic.List[string]
        foreach ($column in $script:dataTable.Columns) {
            [void]$columnNames.Add([string]$column.ColumnName)
        }

        $sortedColumnNames = $columnNames | Sort-Object
        foreach ($columnName in $sortedColumnNames) {
            [void]$clbFilterFields.Items.Add($columnName, $true)
        }

        Update-FilterButtonText
    }
    finally {
        $script:suspendFilterEvents = $false
    }
}

function Get-SelectedFilterColumns {
    $selectedColumns = New-Object System.Collections.Generic.List[string]

    foreach ($checkedItem in $clbFilterFields.CheckedItems) {
        [void]$selectedColumns.Add([string]$checkedItem)
    }

    if ($selectedColumns.Count -eq 0) {
        foreach ($column in $script:dataTable.Columns) {
            [void]$selectedColumns.Add($column.ColumnName)
        }
    }

    return $selectedColumns.ToArray()
}

function Update-FilterSummaryState {
    $filterActive = $false
    if ($script:dataView -and -not [string]::IsNullOrWhiteSpace($script:dataView.RowFilter)) {
        $filterActive = $true
    }

    $chkExportFiltered.Enabled = $filterActive
    if (-not $filterActive) {
        $chkExportFiltered.Checked = $false
    }

    $rowCount = if ($script:dataView) { $script:dataView.Count } else { 0 }
    $totalCount = $script:dataTable.Rows.Count

    if ($totalCount -gt 0 -and $filterActive) {
        Set-Status "Showing $rowCount of $totalCount rows."
    }
    elseif ($totalCount -gt 0) {
        Set-Status "Loaded $totalCount rows successfully." $progress.Maximum $progress.Maximum
    }
    else {
        Set-Status "Ready"
    }
}

function Update-FilterPanelPosition {
    $buttonLocationOnForm = $form.PointToClient($btnFieldPicker.Parent.PointToScreen($btnFieldPicker.Location))
    $x = $buttonLocationOnForm.X
    $y = $buttonLocationOnForm.Y + $btnFieldPicker.Height + 2

    if ($x + $filterPanel.Width -gt ($form.ClientSize.Width - 10)) {
        $x = [Math]::Max(10, ($form.ClientSize.Width - $filterPanel.Width - 10))
    }

    if ($y + $filterPanel.Height -gt ($form.ClientSize.Height - 10)) {
        $y = [Math]::Max($buttonLocationOnForm.Y - $filterPanel.Height - 2, 10)
    }

    $filterPanel.Location = New-Object System.Drawing.Point($x, $y)
    $filterPanel.BringToFront()
}




function Reset-FilterUiState {
    $script:searchDelayTimer.Stop()
    $script:suspendFilterEvents = $true
    try {
        $txtSearch.Text = ""
        $script:highlightSearchText = ""
        $script:highlightColumns = @()
        $btnFieldPicker.Text = "Search Fields ▼"
        $clbFilterFields.Items.Clear()
        $chkExportFiltered.Checked = $false
        $chkExportFiltered.Enabled = $false
        $filterPanel.Visible = $false
    }
    finally {
        $script:suspendFilterEvents = $false
    }
}

function Apply-GridFilter {
    if ($script:suspendFilterEvents -or -not $script:dataView -or -not $script:dataTable -or $script:dataTable.Columns.Count -eq 0) {
        return
    }

    $searchText = $txtSearch.Text.Trim()
    $script:highlightSearchText = $searchText
    $script:highlightColumns = Get-SelectedFilterColumns

    if ([string]::IsNullOrWhiteSpace($searchText)) {
        $script:dataView.RowFilter = ""
        $grid.Refresh()
        Update-FilterSummaryState
        return
    }

    $safeSearchText = Escape-RowFilterValue $searchText
    $selectedColumns = Get-SelectedFilterColumns

    $expressions = foreach ($columnName in $selectedColumns) {
        $safeColumn = $columnName.Replace("]", "]]")
        "CONVERT([$safeColumn], 'System.String') LIKE '%$safeSearchText%'"
    }

    $script:dataView.RowFilter = ($expressions -join " OR ")
    $grid.Refresh()
    Update-FilterSummaryState
}

function Get-ExportRows {
    if ($chkExportFiltered.Checked -and $script:dataView -and -not [string]::IsNullOrWhiteSpace($script:dataView.RowFilter)) {
        $rows = New-Object System.Collections.Generic.List[object]
        for ($i = 0; $i -lt $script:dataView.Count; $i++) {
            [void]$rows.Add($script:dataView[$i])
        }
        return $rows.ToArray()
    }

    $rows = New-Object object[] $script:dataTable.Rows.Count
    for ($i = 0; $i -lt $script:dataTable.Rows.Count; $i++) {
        $rows[$i] = $script:dataTable.Rows[$i]
    }
    return $rows
}

function Get-ExportCellValue {
    param(
        [object]$Row,
        [int]$ColumnIndex
    )

    if ($Row -is [System.Data.DataRowView]) {
        return $Row.Row[$ColumnIndex]
    }

    return $Row[$ColumnIndex]
}

function Get-ColumnWidths {
    param([object[]]$Rows)

    $columnCount = $script:dataTable.Columns.Count
    $widths = New-Object int[] $columnCount

    for ($c = 0; $c -lt $columnCount; $c++) {
        $widths[$c] = $script:dataTable.Columns[$c].ColumnName.Length + 2
    }

    $rowCount = $Rows.Count
    $startTime = Get-Date

    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $Rows[$r]
        for ($c = 0; $c -lt $columnCount; $c++) {
            $len = ([string](Get-ExportCellValue -Row $row -ColumnIndex $c)).Length + 2
            if ($len -gt $widths[$c]) {
                $widths[$c] = $len
            }
        }

        Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total $rowCount -Prefix "Measuring text widths" -Step 250

        if ($script:cancelExport) {
            return $null
        }
    }

    return $widths
}

# ---------------- PARSER ----------------
function Parse-File {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found.", "Error", "OK", "Error")
        return
    }

    $script:suspendFilterEvents = $true
    $filterPanel.Visible = $false
    $script:searchDelayTimer.Stop()
    $grid.SuspendLayout()
    $grid.DataSource = $null
    $script:dataView = $null
    Reset-FilterUiState

    $startTime = Get-Date
    Set-Status "Counting lines..." 0 1

    $totalLines = 0
    $reader = $null

    try {
        $reader = [System.IO.File]::OpenText($FilePath)
        while (($line = $reader.ReadLine()) -ne $null) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $totalLines++
            }
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }

    if ($totalLines -eq 0) {
        $script:dataTable = New-Object System.Data.DataTable
        $grid.ResumeLayout()
        $script:suspendFilterEvents = $false
        Set-Status "No data found." 0 1
        return
    }

    $reader = $null
    $newTable = New-Object System.Data.DataTable

    try {
        $reader = [System.IO.File]::OpenText($FilePath)
        $dataRowsLoaded = 0
        $headers = $null
        $newTable.BeginLoadData()

        while (($line = $reader.ReadLine()) -ne $null) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            if ($null -eq $headers) {
                $headers = $line.Split("`t")
                foreach ($h in $headers) {
                    $columnName = $h.Trim()
                    if ([string]::IsNullOrWhiteSpace($columnName)) {
                        $columnName = "Column$($newTable.Columns.Count + 1)"
                    }

                    $baseName = $columnName
                    $suffix = 1
                    while ($newTable.Columns.Contains($columnName)) {
                        $suffix++
                        $columnName = "$baseName`_$suffix"
                    }

                    [void]$newTable.Columns.Add($columnName)
                }

                Set-Status "Reading rows..." 0 ([Math]::Max(1, $totalLines - 1))
                continue
            }

            $values = Normalize-Row -Values ($line.Split("`t")) -ExpectedCount $headers.Count
            [void]$newTable.Rows.Add($values)
            $dataRowsLoaded++

            Update-ProgressThrottled -StartTime $startTime -Current $dataRowsLoaded -Total ([Math]::Max(1, $totalLines - 1)) -Prefix "Reading" -Step 200
        }

        $newTable.EndLoadData()
        $script:dataTable = $newTable
        $script:dataView = New-Object System.Data.DataView($script:dataTable)
        $grid.DataSource = $script:dataView

        foreach ($column in $grid.Columns) {
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
            if ($column.Width -lt 80) {
                $column.Width = 120
            }
        }

        if ($grid.Columns.Count -gt 0) {
            $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
        }

        Initialize-FilterColumns
        Update-ProgressThrottled -StartTime $startTime -Current $dataRowsLoaded -Total ([Math]::Max(1, $totalLines - 1)) -Prefix "Reading" -Force
        Set-Status "Loaded $dataRowsLoaded rows successfully." $progress.Maximum $progress.Maximum
    }
    catch {
        if ($newTable) {
            try { $newTable.EndLoadData() } catch {}
        }
        $script:dataTable = New-Object System.Data.DataTable
        $script:dataView = $null
        $grid.DataSource = $null
        [System.Windows.Forms.MessageBox]::Show("Failed to parse file:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "Failed to parse file." 0 1
    }
    finally {
        if ($reader) { $reader.Dispose() }
        $script:suspendFilterEvents = $false
        $grid.ResumeLayout()
        Update-FilterSummaryState
    }
}

# ---------------- EVENTS ----------------
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Tab separated text (*.dat;*.txt;*.tsv)|*.dat;*.txt;*.tsv|All files (*.*)|*.*"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dlg.FileName
        Start-ParseIfValid $dlg.FileName
    }
})

$btnProcess.Add_Click({
    Start-ParseIfValid $inputBox.Text
})

$inputBox.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Start-ParseIfValid $inputBox.Text
    }
})

$btnCancel.Add_Click({
    $script:cancelExport = $true
    $lblProgress.Text = "Cancelling export..."
})

$btnFieldPicker.Add_Click({
    if ($clbFilterFields.Items.Count -eq 0) {
        return
    }

    $filterPanel.Visible = -not $filterPanel.Visible
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
})

$btnCloseFilterPanel.Add_Click({
    $filterPanel.Visible = $false
})

$btnSelectAllFields.Add_Click({
    for ($i = 0; $i -lt $clbFilterFields.Items.Count; $i++) {
        $clbFilterFields.SetItemChecked($i, $true)
    }

    Update-FilterButtonText
    Apply-GridFilter
})

$btnSelectNoneFields.Add_Click({
    for ($i = 0; $i -lt $clbFilterFields.Items.Count; $i++) {
        $clbFilterFields.SetItemChecked($i, $false)
    }

    Update-FilterButtonText
    Apply-GridFilter
})

$clbFilterFields.Add_ItemCheck({
    if ($script:suspendFilterEvents) {
        return
    }

    $null = $form.BeginInvoke([System.Windows.Forms.MethodInvoker]{
        if ($script:suspendFilterEvents) { return }
        Update-FilterButtonText
        Apply-GridFilter
    })
})

$btnClearFilter.Add_Click({
    $script:suspendFilterEvents = $true
    try {
        $txtSearch.Text = ""
        for ($i = 0; $i -lt $clbFilterFields.Items.Count; $i++) {
            $clbFilterFields.SetItemChecked($i, $true)
        }
    }
    finally {
        $script:suspendFilterEvents = $false
    }

    Update-FilterButtonText
    Apply-GridFilter
})

$script:searchDelayTimer.Add_Tick({
    $script:searchDelayTimer.Stop()
    if ($script:suspendFilterEvents) { return }
    Apply-GridFilter
})

$txtSearch.Add_TextChanged({
    if ($script:suspendFilterEvents) {
        $script:searchDelayTimer.Stop()
        return
    }

    $script:searchDelayTimer.Stop()
    $script:searchDelayTimer.Start()
})

$form.Add_Resize({
    Update-TopPanelLayout
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
})

$topPanel.Add_Resize({
    Update-TopPanelLayout
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
})

$form.Add_Move({
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
})

$form.Add_MouseDown({
    param($sender, $e)

    if ($filterPanel.Visible) {
        $screenPoint = $form.PointToScreen([System.Drawing.Point]::new($e.X, $e.Y))
        if (-not $filterPanel.RectangleToScreen($filterPanel.ClientRectangle).Contains($screenPoint) -and
            -not $btnFieldPicker.RectangleToScreen($btnFieldPicker.ClientRectangle).Contains($screenPoint)) {
            $filterPanel.Visible = $false
        }
    }
})

$grid.Add_Scroll({
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
})

$grid.Add_CellFormatting({
    param($sender, $e)

    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) {
        return
    }

    $e.CellStyle.BackColor = [System.Drawing.Color]::White
    $e.CellStyle.ForeColor = [System.Drawing.Color]::Black
    $e.CellStyle.SelectionBackColor = [System.Drawing.SystemColors]::Highlight
    $e.CellStyle.SelectionForeColor = [System.Drawing.SystemColors]::HighlightText

    if ([string]::IsNullOrWhiteSpace($script:highlightSearchText)) {
        return
    }

    $columnName = $sender.Columns[$e.ColumnIndex].Name
    if ($script:highlightColumns -and ($script:highlightColumns -notcontains $columnName)) {
        return
    }

    $cellText = ""
    if ($null -ne $e.Value) {
        $cellText = [string]$e.Value
    }

    if ($cellText.IndexOf($script:highlightSearchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
        $e.CellStyle.SelectionBackColor = [System.Drawing.Color]::Goldenrod
        $e.CellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    }
})

# ---------------- TXT EXPORT ----------------
$btnExportTxt.Add_Click({
    if ($script:dataTable.Rows.Count -eq 0 -or $script:dataTable.Columns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "Nothing to export", "OK", "Warning")
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text File (*.txt)|*.txt"

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:cancelExport = $false
    $rowsToExport = Get-ExportRows
    $rowCount = $rowsToExport.Count
    $columnCount = $script:dataTable.Columns.Count
    $startTime = Get-Date

    try {
        Set-Status "Preparing TXT export..." 0 ([Math]::Max(1, $rowCount))
        $colWidths = Get-ColumnWidths -Rows $rowsToExport

        if ($script:cancelExport -or $null -eq $colWidths) {
            Set-Status "Export cancelled." 0 1
            [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
            return
        }

        $writer = [System.IO.StreamWriter]::new($dlg.FileName, $false, [System.Text.Encoding]::UTF8)
        try {
            $headerBuilder = New-Object System.Text.StringBuilder
            [void]$headerBuilder.Append('|')
            for ($c = 0; $c -lt $columnCount; $c++) {
                [void]$headerBuilder.Append($script:dataTable.Columns[$c].ColumnName.PadRight($colWidths[$c]))
                [void]$headerBuilder.Append('|')
            }
            $writer.WriteLine($headerBuilder.ToString())

            for ($r = 0; $r -lt $rowCount; $r++) {
                if ($script:cancelExport) {
                    Set-Status "Export cancelled." 0 1
                    [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                    return
                }

                $row = $rowsToExport[$r]
                $lineBuilder = New-Object System.Text.StringBuilder
                [void]$lineBuilder.Append('|')

                for ($c = 0; $c -lt $columnCount; $c++) {
                    [void]$lineBuilder.Append(([string](Get-ExportCellValue -Row $row -ColumnIndex $c)).PadRight($colWidths[$c]))
                    [void]$lineBuilder.Append('|')
                }

                $writer.WriteLine($lineBuilder.ToString())
                Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total ([Math]::Max(1, $rowCount)) -Prefix "Exporting TXT" -Step 200
            }
        }
        finally {
            if ($writer) { $writer.Dispose() }
        }

        Set-Status "100% complete" ([Math]::Max(1, $rowCount)) ([Math]::Max(1, $rowCount))
        [System.Windows.Forms.MessageBox]::Show("TXT export completed successfully!", "Done", "OK", "Information")
        Update-FilterSummaryState
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("TXT export failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "TXT export failed." 0 1
    }
})

# ---------------- CSV EXPORT ----------------
$btnExportCsv.Add_Click({
    if ($script:dataTable.Rows.Count -eq 0 -or $script:dataTable.Columns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "Nothing to export", "OK", "Warning")
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV File (*.csv)|*.csv"

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:cancelExport = $false
    $rowsToExport = Get-ExportRows
    $rowCount = $rowsToExport.Count
    $columnCount = $script:dataTable.Columns.Count
    $startTime = Get-Date

    try {
        $writer = [System.IO.StreamWriter]::new($dlg.FileName, $false, [System.Text.Encoding]::UTF8)
        try {
            $header = for ($c = 0; $c -lt $columnCount; $c++) {
                Escape-CsvValue $script:dataTable.Columns[$c].ColumnName
            }
            $writer.WriteLine(($header -join ','))

            for ($r = 0; $r -lt $rowCount; $r++) {
                if ($script:cancelExport) {
                    Set-Status "Export cancelled." 0 1
                    [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                    return
                }

                $row = $rowsToExport[$r]
                $values = New-Object string[] $columnCount
                for ($c = 0; $c -lt $columnCount; $c++) {
                    $values[$c] = Escape-CsvValue (Get-ExportCellValue -Row $row -ColumnIndex $c)
                }

                $writer.WriteLine(($values -join ','))
                Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total ([Math]::Max(1, $rowCount)) -Prefix "Exporting CSV" -Step 200
            }
        }
        finally {
            if ($writer) { $writer.Dispose() }
        }

        Set-Status "100% complete" ([Math]::Max(1, $rowCount)) ([Math]::Max(1, $rowCount))
        [System.Windows.Forms.MessageBox]::Show("CSV export completed successfully!", "Done", "OK", "Information")
        Update-FilterSummaryState
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("CSV export failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "CSV export failed." 0 1
    }
})

# ---------------- RUN ----------------
[void]$form.ShowDialog()

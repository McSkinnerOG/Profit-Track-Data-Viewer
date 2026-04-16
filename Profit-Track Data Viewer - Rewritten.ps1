Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- CONFIG ----------------
$Ui = @{
    FormWidth          = 1100
    FormHeight         = 700
    MinWidth           = 900
    MinHeight          = 600
    LeftMargin         = 20
    TopMargin          = 15
    TopPanelWidth      = 1040
    TopPanelHeight     = 95
    GridTop            = 120
    GridHeight         = 440
    ProgressTop        = 575
    StatusTop          = 600
    ExportButtonsTop   = 630
    FilterPanelWidth   = 300
    FilterPanelHeight  = 260
    SearchDelayMs      = 300
    MinColumnWidth     = 120
}

# ---------------- STORAGE ----------------
$script:dataTable = New-Object System.Data.DataTable
$script:dataView = $null
$script:cancelExport = $false
$script:highlightSearchText = ""
$script:highlightColumns = @()
$script:suspendFilterEvents = $false
$script:searchDelayTimer = New-Object System.Windows.Forms.Timer
$script:searchDelayTimer.Interval = $Ui.SearchDelayMs

# ---------------- HELPERS ----------------
function Show-Info {
    param([string]$Text, [string]$Title = "Information")
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, "OK", "Information")
}

function Show-Warning {
    param([string]$Text, [string]$Title = "Warning")
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, "OK", "Warning")
}

function Show-Error {
    param([string]$Text, [string]$Title = "Error")
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, "OK", "Error")
}

function New-UiLabel {
    param(
        [string]$Text,
        [System.Windows.Forms.Padding]$Margin = ([System.Windows.Forms.Padding]::new(0,0,10,0))
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Dock = [System.Windows.Forms.DockStyle]::Fill
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.Margin = $Margin
    return $label
}

function New-UiButton {
    param(
        [string]$Text,
        [System.Windows.Forms.Padding]$Margin = ([System.Windows.Forms.Padding]::new(0)),
        [System.Windows.Forms.DockStyle]$Dock = [System.Windows.Forms.DockStyle]::Fill
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Dock = $Dock
    $button.Margin = $Margin
    return $button
}

function New-UiTableLayout {
    param(
        [int]$Columns,
        [int]$Rows,
        [System.Windows.Forms.Padding]$Margin = ([System.Windows.Forms.Padding]::new(0)),
        [System.Windows.Forms.Padding]$Padding = ([System.Windows.Forms.Padding]::new(0))
    )

    $panel = New-Object System.Windows.Forms.TableLayoutPanel
    $panel.ColumnCount = $Columns
    $panel.RowCount = $Rows
    $panel.Margin = $Margin
    $panel.Padding = $Padding
    $panel.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
    return $panel
}

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

function Get-SafeCount {
    param([int]$Value)
    return [Math]::Max(1, $Value)
}

function Get-ETA {
    param(
        [datetime]$StartTime,
        [int]$Current,
        [int]$Total
    )

    if ($Current -le 0 -or $Total -le 0) { return "Calculating..." }

    $elapsed = (Get-Date) - $StartTime
    if ($elapsed.TotalSeconds -le 0) { return "Calculating..." }

    $rate = $Current / $elapsed.TotalSeconds
    if ($rate -le 0) { return "Calculating..." }

    $remaining = ($Total - $Current) / $rate
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

    $percent = if ($Total -gt 0) { [Math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    $eta = Get-ETA -StartTime $StartTime -Current $Current -Total $Total
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

    $result = New-Object string[] $ExpectedCount
    for ($i = 0; $i -lt $ExpectedCount; $i++) {
        if ($i -lt $Values.Count) {
            $result[$i] = [string]$Values[$i].Trim()
        }
        else {
            $result[$i] = ""
        }
    }

    return $result
}

function Escape-RowFilterValue {
    param([string]$Value)

    if ($null -eq $Value) { return "" }

    $escaped = $Value.Replace("'", "''")
    $escaped = $escaped.Replace("[", "[[")
    $escaped = $escaped.Replace("]", "]]" )
    $escaped = $escaped.Replace("%", "[%]")
    $escaped = $escaped.Replace("*", "[*]")
    return $escaped
}

function Get-UniqueColumnName {
    param(
        [System.Data.DataTable]$Table,
        [string]$ColumnName
    )

    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        $ColumnName = "Column$($Table.Columns.Count + 1)"
    }

    $baseName = $ColumnName.Trim()
    $name = $baseName
    $suffix = 1

    while ($Table.Columns.Contains($name)) {
        $suffix++
        $name = "${baseName}_$suffix"
    }

    return $name
}

function Set-AllFilterFieldChecks {
    param([bool]$Checked)

    for ($i = 0; $i -lt $clbFilterFields.Items.Count; $i++) {
        $clbFilterFields.SetItemChecked($i, $Checked)
    }
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
        $sortedColumnNames = @(
            $script:dataTable.Columns |
            ForEach-Object { [string]$_.ColumnName } |
            Sort-Object
        )

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
            [void]$selectedColumns.Add([string]$column.ColumnName)
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
        $x = [Math]::Max(10, $form.ClientSize.Width - $filterPanel.Width - 10)
    }

    if ($y + $filterPanel.Height -gt ($form.ClientSize.Height - 10)) {
        $y = [Math]::Max($buttonLocationOnForm.Y - $filterPanel.Height - 2, 10)
    }

    $filterPanel.Location = New-Object System.Drawing.Point($x, $y)
    $filterPanel.BringToFront()
}

function Update-FilterPanelIfVisible {
    if ($filterPanel.Visible) {
        Update-FilterPanelPosition
    }
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

function Reset-GridState {
    $filterPanel.Visible = $false
    $script:searchDelayTimer.Stop()
    $grid.SuspendLayout()
    $grid.DataSource = $null
    $script:dataView = $null
    Reset-FilterUiState
}

function Get-NonEmptyLineCount {
    param([string]$FilePath)

    $count = 0
    $reader = $null
    try {
        $reader = [System.IO.File]::OpenText($FilePath)
        while (($line = $reader.ReadLine()) -ne $null) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $count++
            }
        }
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }

    return $count
}

function New-DataTableFromFile {
    param(
        [string]$FilePath,
        [int]$TotalLines,
        [datetime]$StartTime
    )

    $reader = $null
    $table = New-Object System.Data.DataTable
    $dataRowsLoaded = 0
    $headers = $null

    try {
        $reader = [System.IO.File]::OpenText($FilePath)
        $table.BeginLoadData()

        while (($line = $reader.ReadLine()) -ne $null) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            if ($null -eq $headers) {
                $headers = $line.Split("`t")
                foreach ($header in $headers) {
                    [void]$table.Columns.Add((Get-UniqueColumnName -Table $table -ColumnName $header))
                }

                Set-Status "Reading rows..." 0 (Get-SafeCount ($TotalLines - 1))
                continue
            }

            $values = Normalize-Row -Values ($line.Split("`t")) -ExpectedCount $headers.Count
            [void]$table.Rows.Add($values)
            $dataRowsLoaded++

            Update-ProgressThrottled -StartTime $StartTime -Current $dataRowsLoaded -Total (Get-SafeCount ($TotalLines - 1)) -Prefix "Reading" -Step 200
        }

        $table.EndLoadData()
        return [pscustomobject]@{
            Table = $table
            DataRowsLoaded = $dataRowsLoaded
        }
    }
    catch {
        try { $table.EndLoadData() } catch {}
        throw
    }
    finally {
        if ($reader) { $reader.Dispose() }
    }
}

function Bind-DataTableToGrid {
    $script:dataView = New-Object System.Data.DataView($script:dataTable)
    $grid.DataSource = $script:dataView

    foreach ($column in $grid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        if ($column.Width -lt 80) {
            $column.Width = $Ui.MinColumnWidth
        }
    }

    if ($grid.Columns.Count -gt 0) {
        $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
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
        $safeColumn = $columnName.Replace("]", "]]" )
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

        Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total (Get-SafeCount $rowCount) -Prefix "Measuring text widths" -Step 250

        if ($script:cancelExport) {
            return $null
        }
    }

    return $widths
}

function Invoke-CancelledExport {
    Set-Status "Export cancelled." 0 1
    Show-Info "Export cancelled."
}

function Export-TxtHeader {
    param([System.IO.StreamWriter]$Writer, [int]$ColumnCount, [int[]]$ColumnWidths)

    $headerBuilder = New-Object System.Text.StringBuilder
    [void]$headerBuilder.Append('|')
    for ($c = 0; $c -lt $ColumnCount; $c++) {
        [void]$headerBuilder.Append($script:dataTable.Columns[$c].ColumnName.PadRight($ColumnWidths[$c]))
        [void]$headerBuilder.Append('|')
    }
    $Writer.WriteLine($headerBuilder.ToString())
}

function Export-TxtRow {
    param([System.IO.StreamWriter]$Writer, [object]$Row, [int]$ColumnCount, [int[]]$ColumnWidths)

    $lineBuilder = New-Object System.Text.StringBuilder
    [void]$lineBuilder.Append('|')
    for ($c = 0; $c -lt $ColumnCount; $c++) {
        [void]$lineBuilder.Append(([string](Get-ExportCellValue -Row $Row -ColumnIndex $c)).PadRight($ColumnWidths[$c]))
        [void]$lineBuilder.Append('|')
    }
    $Writer.WriteLine($lineBuilder.ToString())
}

function Export-CsvHeader {
    param([System.IO.StreamWriter]$Writer, [int]$ColumnCount)

    $header = for ($c = 0; $c -lt $ColumnCount; $c++) {
        Escape-CsvValue $script:dataTable.Columns[$c].ColumnName
    }
    $Writer.WriteLine(($header -join ','))
}

function Export-CsvRow {
    param([System.IO.StreamWriter]$Writer, [object]$Row, [int]$ColumnCount)

    $values = New-Object string[] $ColumnCount
    for ($c = 0; $c -lt $ColumnCount; $c++) {
        $values[$c] = Escape-CsvValue (Get-ExportCellValue -Row $Row -ColumnIndex $c)
    }
    $Writer.WriteLine(($values -join ','))
}

function Invoke-DataExport {
    param(
        [string]$ExportName,
        [string]$DialogFilter,
        [scriptblock]$WriteHeader,
        [scriptblock]$WriteRow,
        [switch]$NeedsColumnWidths
    )

    if ($script:dataTable.Rows.Count -eq 0 -or $script:dataTable.Columns.Count -eq 0) {
        Show-Warning "There is no data to export." "Nothing to export"
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = $DialogFilter
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:cancelExport = $false
    $rowsToExport = Get-ExportRows
    $rowCount = $rowsToExport.Count
    $columnCount = $script:dataTable.Columns.Count
    $safeRowCount = Get-SafeCount $rowCount
    $startTime = Get-Date
    $columnWidths = $null

    try {
        if ($NeedsColumnWidths) {
            Set-Status "Preparing $ExportName export..." 0 $safeRowCount
            $columnWidths = Get-ColumnWidths -Rows $rowsToExport

            if ($script:cancelExport -or $null -eq $columnWidths) {
                Invoke-CancelledExport
                return
            }
        }

        $writer = [System.IO.StreamWriter]::new($dialog.FileName, $false, [System.Text.Encoding]::UTF8)
        try {
            & $WriteHeader $writer $columnCount $columnWidths

            for ($r = 0; $r -lt $rowCount; $r++) {
                if ($script:cancelExport) {
                    Invoke-CancelledExport
                    return
                }

                & $WriteRow $writer $rowsToExport[$r] $columnCount $columnWidths
                Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total $safeRowCount -Prefix "Exporting $ExportName" -Step 200
            }
        }
        finally {
            if ($writer) { $writer.Dispose() }
        }

        Set-Status "100% complete" $safeRowCount $safeRowCount
        Show-Info "$ExportName export completed successfully!" "Done"
        Update-FilterSummaryState
    }
    catch {
        Show-Error "$ExportName export failed:`r`n$($_.Exception.Message)"
        Set-Status "$ExportName export failed." 0 1
    }
}

function Parse-File {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) {
        Show-Error "File not found."
        return
    }

    $script:suspendFilterEvents = $true
    Reset-GridState
    $startTime = Get-Date
    Set-Status "Counting lines..." 0 1

    try {
        $totalLines = Get-NonEmptyLineCount -FilePath $FilePath
        if ($totalLines -eq 0) {
            $script:dataTable = New-Object System.Data.DataTable
            Set-Status "No data found." 0 1
            return
        }

        $result = New-DataTableFromFile -FilePath $FilePath -TotalLines $totalLines -StartTime $startTime
        $script:dataTable = $result.Table
        Bind-DataTableToGrid
        Initialize-FilterColumns
        Update-ProgressThrottled -StartTime $startTime -Current $result.DataRowsLoaded -Total (Get-SafeCount ($totalLines - 1)) -Prefix "Reading" -Force
        Set-Status "Loaded $($result.DataRowsLoaded) rows successfully." $progress.Maximum $progress.Maximum
    }
    catch {
        $script:dataTable = New-Object System.Data.DataTable
        $script:dataView = $null
        $grid.DataSource = $null
        Show-Error "Failed to parse file:`r`n$($_.Exception.Message)"
        Set-Status "Failed to parse file." 0 1
    }
    finally {
        $script:suspendFilterEvents = $false
        $grid.ResumeLayout()
        Update-FilterSummaryState
    }
}

function Start-ParseIfValid {
    param([string]$FilePath)

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        return
    }

    if (Test-Path -LiteralPath $FilePath) {
        Parse-File -FilePath $FilePath
    }
    else {
        Show-Warning "Please choose a valid file first." "Missing file"
    }
}

# ---------------- MAIN FORM ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Profit-Track Data Formatter ~ By James Connolly - 2026"
$form.Size = New-Object System.Drawing.Size($Ui.FormWidth, $Ui.FormHeight)
$form.MinimumSize = New-Object System.Drawing.Size($Ui.MinWidth, $Ui.MinHeight)
$form.StartPosition = "CenterScreen"

# ---------------- TOP PANEL ----------------
$topPanel = New-UiTableLayout -Columns 1 -Rows 2
$topPanel.Location = New-Object System.Drawing.Point($Ui.LeftMargin, $Ui.TopMargin)
$topPanel.Size = New-Object System.Drawing.Size($Ui.TopPanelWidth, $Ui.TopPanelHeight)
$topPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$topPanel.AutoSize = $false
[void]$topPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
[void]$topPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
$form.Controls.Add($topPanel)

# ---------------- FILE INPUT ROW ----------------
$fileRow = New-UiTableLayout -Columns 4 -Rows 1 -Margin ([System.Windows.Forms.Padding]::new(0,0,0,6))
$fileRow.Dock = [System.Windows.Forms.DockStyle]::Fill
[void]$fileRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
[void]$fileRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$fileRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
[void]$fileRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 1)))
$topPanel.Controls.Add($fileRow, 0, 0)

$lblFile = New-UiLabel -Text "File:"
$fileRow.Controls.Add($lblFile, 0, 0)

$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$inputBox.Margin = [System.Windows.Forms.Padding]::new(0,2,10,0)
$fileRow.Controls.Add($inputBox, 1, 0)

$btnBrowse = New-UiButton -Text "Browse..."
$fileRow.Controls.Add($btnBrowse, 2, 0)

$btnProcess = New-UiButton -Text "Process"
$btnProcess.Visible = $false
$fileRow.Controls.Add($btnProcess, 3, 0)

# ---------------- SEARCH ROW ----------------
$searchRow = New-UiTableLayout -Columns 5 -Rows 1
$searchRow.Dock = [System.Windows.Forms.DockStyle]::Fill
[void]$searchRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
[void]$searchRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$searchRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170)))
[void]$searchRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 115)))
[void]$searchRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 210)))
$topPanel.Controls.Add($searchRow, 0, 1)

$lblSearch = New-UiLabel -Text "Search:"
$searchRow.Controls.Add($lblSearch, 0, 0)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtSearch.Margin = [System.Windows.Forms.Padding]::new(0,2,10,0)
$searchRow.Controls.Add($txtSearch, 1, 0)

$btnFieldPicker = New-UiButton -Text "Fields ▼" -Margin ([System.Windows.Forms.Padding]::new(0,0,10,0))
$searchRow.Controls.Add($btnFieldPicker, 2, 0)

$btnClearFilter = New-UiButton -Text "Clear" -Margin ([System.Windows.Forms.Padding]::new(0,0,10,0))
$searchRow.Controls.Add($btnClearFilter, 3, 0)

$chkExportFiltered = New-Object System.Windows.Forms.CheckBox
$chkExportFiltered.Text = "Export filtered rows only"
$chkExportFiltered.AutoSize = $true
$chkExportFiltered.Dock = [System.Windows.Forms.DockStyle]::Fill
$chkExportFiltered.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$chkExportFiltered.Margin = [System.Windows.Forms.Padding]::new(0,4,0,0)
$chkExportFiltered.Enabled = $false
$searchRow.Controls.Add($chkExportFiltered, 4, 0)

# ---------------- FILTER POPUP PANEL ----------------
$filterPanel = New-Object System.Windows.Forms.Panel
$filterPanel.Size = New-Object System.Drawing.Size($Ui.FilterPanelWidth, $Ui.FilterPanelHeight)
$filterPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$filterPanel.Visible = $false
$filterPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($filterPanel)
$filterPanel.BringToFront()

$lblFilterPanel = New-Object System.Windows.Forms.Label
$lblFilterPanel.Text = "Choose one or more fields to search:"
$lblFilterPanel.AutoSize = $true
$lblFilterPanel.Location = New-Object System.Drawing.Point(8, 8)
$filterPanel.Controls.Add($lblFilterPanel)

$clbFilterFields = New-Object System.Windows.Forms.CheckedListBox
$clbFilterFields.Location = New-Object System.Drawing.Point(8, 30)
$clbFilterFields.Size = New-Object System.Drawing.Size(282, 185)
$clbFilterFields.CheckOnClick = $true
$clbFilterFields.IntegralHeight = $false
$clbFilterFields.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$filterPanel.Controls.Add($clbFilterFields)

$btnSelectAllFields = New-Object System.Windows.Forms.Button
$btnSelectAllFields.Text = "All"
$btnSelectAllFields.Size = New-Object System.Drawing.Size(55, 25)
$btnSelectAllFields.Location = New-Object System.Drawing.Point(8, 225)
$btnSelectAllFields.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$filterPanel.Controls.Add($btnSelectAllFields)

$btnSelectNoneFields = New-Object System.Windows.Forms.Button
$btnSelectNoneFields.Text = "None"
$btnSelectNoneFields.Size = New-Object System.Drawing.Size(55, 25)
$btnSelectNoneFields.Location = New-Object System.Drawing.Point(70, 225)
$btnSelectNoneFields.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$filterPanel.Controls.Add($btnSelectNoneFields)

$btnCloseFilterPanel = New-Object System.Windows.Forms.Button
$btnCloseFilterPanel.Text = "Close"
$btnCloseFilterPanel.Size = New-Object System.Drawing.Size(60, 25)
$btnCloseFilterPanel.Location = New-Object System.Drawing.Point(230, 225)
$btnCloseFilterPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$filterPanel.Controls.Add($btnCloseFilterPanel)

# ---------------- GRID ----------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point($Ui.LeftMargin, $Ui.GridTop)
$grid.Size = New-Object System.Drawing.Size($Ui.TopPanelWidth, $Ui.GridHeight)
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
$progress.Location = New-Object System.Drawing.Point($Ui.LeftMargin, $Ui.ProgressTop)
$progress.Size = New-Object System.Drawing.Size($Ui.TopPanelWidth, 20)
$progress.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progress)

# ---------------- STATUS LABEL ----------------
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point($Ui.LeftMargin, $Ui.StatusTop)
$lblProgress.Size = New-Object System.Drawing.Size(760, 20)
$lblProgress.Text = "Ready"
$lblProgress.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($lblProgress)

# ---------------- EXPORT BUTTONS ----------------
$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Text = "Export TXT"
$btnExportTxt.Size = New-Object System.Drawing.Size(95, 28)
$btnExportTxt.Location = New-Object System.Drawing.Point($Ui.LeftMargin, $Ui.ExportButtonsTop)
$btnExportTxt.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnExportTxt)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Size = New-Object System.Drawing.Size(95, 28)
$btnExportCsv.Location = New-Object System.Drawing.Point(125, $Ui.ExportButtonsTop)
$btnExportCsv.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($btnExportCsv)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Export"
$btnCancel.Size = New-Object System.Drawing.Size(110, 28)
$btnCancel.Location = New-Object System.Drawing.Point(235, $Ui.ExportButtonsTop)
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
    Start-ParseIfValid -FilePath $file
})

# ---------------- EVENTS ----------------
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Tab separated text (*.dat;*.txt;*.tsv)|*.dat;*.txt;*.tsv|All files (*.*)|*.*"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dialog.FileName
        Start-ParseIfValid -FilePath $dialog.FileName
    }
})

$btnProcess.Add_Click({
    Start-ParseIfValid -FilePath $inputBox.Text
})

$inputBox.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Start-ParseIfValid -FilePath $inputBox.Text
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
    Update-FilterPanelIfVisible
})

$btnCloseFilterPanel.Add_Click({
    $filterPanel.Visible = $false
})

$btnSelectAllFields.Add_Click({
    Set-AllFilterFieldChecks -Checked $true
    Update-FilterButtonText
    Apply-GridFilter
})

$btnSelectNoneFields.Add_Click({
    Set-AllFilterFieldChecks -Checked $false
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
        Set-AllFilterFieldChecks -Checked $true
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

$form.Add_Resize({ Update-FilterPanelIfVisible })
$topPanel.Add_Resize({ Update-FilterPanelIfVisible })
$form.Add_Move({ Update-FilterPanelIfVisible })
$grid.Add_Scroll({ Update-FilterPanelIfVisible })

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

    $cellText = if ($null -ne $e.Value) { [string]$e.Value } else { "" }
    if ($cellText.IndexOf($script:highlightSearchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
        $e.CellStyle.SelectionBackColor = [System.Drawing.Color]::Goldenrod
        $e.CellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    }
})

$btnExportTxt.Add_Click({
    Invoke-DataExport -ExportName "TXT" -DialogFilter "Text File (*.txt)|*.txt" -WriteHeader ${function:Export-TxtHeader} -WriteRow ${function:Export-TxtRow} -NeedsColumnWidths
})

$btnExportCsv.Add_Click({
    Invoke-DataExport -ExportName "CSV" -DialogFilter "CSV File (*.csv)|*.csv" -WriteHeader ${function:Export-CsvHeader} -WriteRow ${function:Export-CsvRow}
})

# ---------------- RUN ----------------
[void]$form.ShowDialog()

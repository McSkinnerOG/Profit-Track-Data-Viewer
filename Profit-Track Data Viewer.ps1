Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- MAIN FORM ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Profit-Track Data Formatter ~ By James Connolly - 2026"
$form.Size = New-Object System.Drawing.Size(1000,650)
$form.StartPosition = "CenterScreen"

# ---------------- INPUT BOX ----------------
$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point(20,20)
$inputBox.Size = New-Object System.Drawing.Size(700,25)
$inputBox.Anchor = "Top,Left,Right"
$form.Controls.Add($inputBox)

# ---------------- BUTTONS ----------------
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(740,18)
$btnBrowse.Anchor = "Top,Right"
$form.Controls.Add($btnBrowse)

$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Process"
$btnProcess.Location = New-Object System.Drawing.Point(820,18)
$btnProcess.Anchor = "Top,Right"
$form.Controls.Add($btnProcess)

# ---------------- GRID ----------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20,60)
$grid.Size = New-Object System.Drawing.Size(940,460)
$grid.Anchor = "Top,Bottom,Left,Right"
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
$progress.Location = New-Object System.Drawing.Point(20,530)
$progress.Size = New-Object System.Drawing.Size(940,20)
$progress.Anchor = "Bottom,Left,Right"
$form.Controls.Add($progress)

# ---------------- LABEL ----------------
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(20,555)
$lblProgress.Size = New-Object System.Drawing.Size(700,20)
$lblProgress.Text = "Ready"
$lblProgress.Anchor = "Bottom,Left,Right"
$form.Controls.Add($lblProgress)

# ---------------- CANCEL BUTTON ----------------
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Export"
$btnCancel.Location = New-Object System.Drawing.Point(220,585)
$btnCancel.Anchor = "Bottom,Left"
$form.Controls.Add($btnCancel)

# ---------------- EXPORT BUTTONS ----------------
$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Text = "Export TXT"
$btnExportTxt.Location = New-Object System.Drawing.Point(20,585)
$btnExportTxt.Anchor = "Bottom,Left"
$form.Controls.Add($btnExportTxt)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Location = New-Object System.Drawing.Point(120,585)
$btnExportCsv.Anchor = "Bottom,Left"
$form.Controls.Add($btnExportCsv)

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
})

# ---------------- STORAGE ----------------
$dataTable = New-Object System.Data.DataTable
$script:cancelExport = $false

# ---------------- HELPERS ----------------
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

# ---------------- PARSER ----------------
function Parse-File {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found.", "Error", "OK", "Error")
        return
    }

    $grid.SuspendLayout()
    $grid.DataSource = $null

    $dataTable.Clear()
    $dataTable.Columns.Clear()

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
        $grid.ResumeLayout()
        Set-Status "No data found." 0 1
        return
    }

    $reader = $null

    try {
        $reader = [System.IO.File]::OpenText($FilePath)
        $lineNumber = 0
        $dataRowsLoaded = 0
        $headers = $null
        $dataTable.BeginLoadData()

        while (($line = $reader.ReadLine()) -ne $null) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $lineNumber++

            if ($null -eq $headers) {
                $headers = $line.Split("`t")
                foreach ($h in $headers) {
                    $columnName = $h.Trim()
                    if ([string]::IsNullOrWhiteSpace($columnName)) {
                        $columnName = "Column$($dataTable.Columns.Count + 1)"
                    }

                    $baseName = $columnName
                    $suffix = 1
                    while ($dataTable.Columns.Contains($columnName)) {
                        $suffix++
                        $columnName = "$baseName`_$suffix"
                    }

                    [void]$dataTable.Columns.Add($columnName)
                }

                Set-Status "Reading rows..." 0 ($totalLines - 1)
                continue
            }

            $values = Normalize-Row -Values ($line.Split("`t")) -ExpectedCount $headers.Count
            [void]$dataTable.Rows.Add($values)
            $dataRowsLoaded++

            Update-ProgressThrottled -StartTime $startTime -Current $dataRowsLoaded -Total ($totalLines - 1) -Prefix "Reading" -Step 200
        }

        $dataTable.EndLoadData()
        $grid.DataSource = $dataTable

        foreach ($column in $grid.Columns) {
            $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
            if ($column.Width -lt 80) {
                $column.Width = 120
            }
        }

        if ($grid.Columns.Count -gt 0) {
            $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
        }

        Update-ProgressThrottled -StartTime $startTime -Current $dataRowsLoaded -Total ([Math]::Max(1, $totalLines - 1)) -Prefix "Reading" -Force
        Set-Status "Loaded $dataRowsLoaded rows successfully." $progress.Maximum $progress.Maximum
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to parse file:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "Failed to parse file." 0 1
    }
    finally {
        if ($reader) { $reader.Dispose() }
        $grid.ResumeLayout()
    }
}

function Get-ColumnWidths {
    $columnCount = $dataTable.Columns.Count
    $widths = New-Object int[] $columnCount

    for ($c = 0; $c -lt $columnCount; $c++) {
        $widths[$c] = $dataTable.Columns[$c].ColumnName.Length + 2
    }

    $rowCount = $dataTable.Rows.Count
    $startTime = Get-Date

    for ($r = 0; $r -lt $rowCount; $r++) {
        $row = $dataTable.Rows[$r]
        for ($c = 0; $c -lt $columnCount; $c++) {
            $len = ([string]$row[$c]).Length + 2
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

# ---------------- EVENTS ----------------
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Tab separated text (*.dat;*.txt;*.tsv)|*.dat;*.txt;*.tsv|All files (*.*)|*.*"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputBox.Text = $dlg.FileName
    }
})

$btnProcess.Add_Click({
    if (Test-Path -LiteralPath $inputBox.Text) {
        Parse-File $inputBox.Text
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please choose a valid file first.", "Missing file", "OK", "Warning")
    }
})

# ---------------- CANCEL ----------------
$btnCancel.Add_Click({
    $script:cancelExport = $true
    $lblProgress.Text = "Cancelling export..."
})

# ---------------- TXT EXPORT ----------------
$btnExportTxt.Add_Click({
    if ($dataTable.Rows.Count -eq 0 -or $dataTable.Columns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "Nothing to export", "OK", "Warning")
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text File (*.txt)|*.txt"

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:cancelExport = $false
    $rowCount = $dataTable.Rows.Count
    $columnCount = $dataTable.Columns.Count
    $startTime = Get-Date

    try {
        Set-Status "Preparing TXT export..." 0 ([Math]::Max(1, $rowCount))
        $colWidths = Get-ColumnWidths

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
                [void]$headerBuilder.Append($dataTable.Columns[$c].ColumnName.PadRight($colWidths[$c]))
                [void]$headerBuilder.Append('|')
            }
            $writer.WriteLine($headerBuilder.ToString())

            for ($r = 0; $r -lt $rowCount; $r++) {
                if ($script:cancelExport) {
                    Set-Status "Export cancelled." 0 1
                    [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                    return
                }

                $row = $dataTable.Rows[$r]
                $lineBuilder = New-Object System.Text.StringBuilder
                [void]$lineBuilder.Append('|')

                for ($c = 0; $c -lt $columnCount; $c++) {
                    [void]$lineBuilder.Append(([string]$row[$c]).PadRight($colWidths[$c]))
                    [void]$lineBuilder.Append('|')
                }

                $writer.WriteLine($lineBuilder.ToString())
                Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total $rowCount -Prefix "Exporting TXT" -Step 200
            }
        }
        finally {
            if ($writer) { $writer.Dispose() }
        }

        Set-Status "100% complete" $rowCount $rowCount
        [System.Windows.Forms.MessageBox]::Show("TXT export completed successfully!", "Done", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("TXT export failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "TXT export failed." 0 1
    }
})

# ---------------- CSV EXPORT ----------------
$btnExportCsv.Add_Click({
    if ($dataTable.Rows.Count -eq 0 -or $dataTable.Columns.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There is no data to export.", "Nothing to export", "OK", "Warning")
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV File (*.csv)|*.csv"

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $script:cancelExport = $false
    $rowCount = $dataTable.Rows.Count
    $columnCount = $dataTable.Columns.Count
    $startTime = Get-Date

    try {
        $writer = [System.IO.StreamWriter]::new($dlg.FileName, $false, [System.Text.Encoding]::UTF8)
        try {
            $header = for ($c = 0; $c -lt $columnCount; $c++) {
                Escape-CsvValue $dataTable.Columns[$c].ColumnName
            }
            $writer.WriteLine(($header -join ','))

            for ($r = 0; $r -lt $rowCount; $r++) {
                if ($script:cancelExport) {
                    Set-Status "Export cancelled." 0 1
                    [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                    return
                }

                $row = $dataTable.Rows[$r]
                $values = New-Object string[] $columnCount
                for ($c = 0; $c -lt $columnCount; $c++) {
                    $values[$c] = Escape-CsvValue $row[$c]
                }

                $writer.WriteLine(($values -join ','))
                Update-ProgressThrottled -StartTime $startTime -Current ($r + 1) -Total $rowCount -Prefix "Exporting CSV" -Step 200
            }
        }
        finally {
            if ($writer) { $writer.Dispose() }
        }

        Set-Status "100% complete" $rowCount $rowCount
        [System.Windows.Forms.MessageBox]::Show("CSV export completed successfully!", "Done", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("CSV export failed:`r`n$($_.Exception.Message)", "Error", "OK", "Error")
        Set-Status "CSV export failed." 0 1
    }
})

# ---------------- RUN ----------------
[void]$form.ShowDialog()

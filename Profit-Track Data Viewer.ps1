Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- MAIN FORM ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Profit-Track Data Formatter ~ By James Connolly - 2026"
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

# ---------------- INPUT BOX ----------------
$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point(20,20)
$inputBox.Size = New-Object System.Drawing.Size(600,25)
$inputBox.Anchor = "Top,Left,Right"
$form.Controls.Add($inputBox)

# ---------------- BUTTONS ----------------
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse"
$btnBrowse.Location = New-Object System.Drawing.Point(640,18)
$btnBrowse.Anchor = "Top,Right"
$form.Controls.Add($btnBrowse)

$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Process"
$btnProcess.Location = New-Object System.Drawing.Point(720,18)
$btnProcess.Anchor = "Top,Right"
$form.Controls.Add($btnProcess)

# ---------------- GRID ----------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20,60)
$grid.Size = New-Object System.Drawing.Size(840,400)
$grid.AutoSizeColumnsMode = "DisplayedCells"
$grid.AutoSizeRowsMode = "AllCells"
$grid.ScrollBars = "Both" 
$grid.Anchor = "Top,Bottom,Left,Right"
$form.Controls.Add($grid)

# ---------------- PROGRESS BAR ----------------
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20,470)
$progress.Size = New-Object System.Drawing.Size(840,20)
$progress.Anchor = "Bottom,Left,Right"
$form.Controls.Add($progress)

# ---------------- LABEL ----------------
$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point(20,490)
$lblProgress.Size = New-Object System.Drawing.Size(600,20)
$lblProgress.Text = "Ready"
$lblProgress.Anchor = "Bottom,Left"
$form.Controls.Add($lblProgress)

# ---------------- CANCEL BUTTON ----------------
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel Export"
$btnCancel.Location = New-Object System.Drawing.Point(220,520)
$btnCancel.Anchor = "Bottom,Left"
$form.Controls.Add($btnCancel)

# ---------------- EXPORT BUTTONS ----------------
$btnExportTxt = New-Object System.Windows.Forms.Button
$btnExportTxt.Text = "Export TXT"
$btnExportTxt.Location = New-Object System.Drawing.Point(20,520)
$btnExportTxt.Anchor = "Bottom,Left"
$form.Controls.Add($btnExportTxt)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Location = New-Object System.Drawing.Point(120,520)
$btnExportCsv.Anchor = "Bottom,Left"
$form.Controls.Add($btnExportCsv)

# ---------------- DRAG & DROP ----------------
$form.AllowDrop = $true
$form.Add_DragEnter({
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $_.Effect = "Copy"
    }
})

$form.Add_DragDrop({
    $file = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)[0]
    $inputBox.Text = $file
})

# ---------------- STORAGE ----------------
$dataTable = New-Object System.Data.DataTable
$script:cancelExport = $false

# ---------------- ETA FUNCTION ----------------
function Get-ETA {
    param($startTime, $current, $total)

    if ($current -eq 0) { return "Calculating..." }

    $elapsed = (Get-Date) - $startTime
    $rate = $current / $elapsed.TotalSeconds
    $remaining = ($total - $current) / $rate

    return ([TimeSpan]::FromSeconds($remaining)).ToString("hh\:mm\:ss")
}

# ---------------- PARSER ----------------
function Parse-File($filePath) {

    $dataTable.Clear()
    $dataTable.Columns.Clear()

    $lines = Get-Content $filePath | Where-Object { $_.Trim() -ne "" }

    $headers = $lines[0] -split "`t"

    foreach ($h in $headers) {
        $dataTable.Columns.Add($h.Trim()) | Out-Null
    }

    $progress.Maximum = $lines.Count
    $progress.Value = 0

    for ($i = 1; $i -lt $lines.Count; $i++) {

        $row = $lines[$i] -split "`t"
        $row = $row | ForEach-Object { $_.Trim() }

        while ($row.Count -lt $headers.Count) {
            $row += ""
        }

        if ($row.Count -gt $headers.Count) {
            $row = $row[0..($headers.Count-1)]
        }

        $dataTable.Rows.Add($row) | Out-Null

        $progress.Value = $i
        [System.Windows.Forms.Application]::DoEvents()
    }

    $grid.DataSource = $dataTable
    $grid.AutoResizeColumns()
    $grid.AutoResizeRows()
}

# ---------------- EVENTS ----------------
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    if ($dlg.ShowDialog() -eq "OK") {
        $inputBox.Text = $dlg.FileName
    }
})

$btnProcess.Add_Click({
    if (Test-Path $inputBox.Text) {
        Parse-File $inputBox.Text
    }
})

# ---------------- CANCEL ----------------
$btnCancel.Add_Click({
    $script:cancelExport = $true
})

# ---------------- TXT EXPORT (FULL FEATURED) ----------------
$btnExportTxt.Add_Click({

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text File (*.txt)|*.txt"

    if ($dlg.ShowDialog() -eq "OK") {

        $script:cancelExport = $false
        $startTime = Get-Date

        $progress.Value = 0
        $progress.Maximum = $dataTable.Rows.Count

        # ---- column widths ----
        $colWidths = @()

        for ($c = 0; $c -lt $dataTable.Columns.Count; $c++) {

            $maxLen = $dataTable.Columns[$c].ColumnName.Length

            foreach ($row in $dataTable.Rows) {
                $val = [string]$row[$c]
                if ($val.Length -gt $maxLen) { $maxLen = $val.Length }
            }

            $colWidths += $maxLen + 2
        }

        # ---- header ----
        $out = "|"
        for ($c = 0; $c -lt $dataTable.Columns.Count; $c++) {
            $out += ($dataTable.Columns[$c].ColumnName.PadRight($colWidths[$c])) + "|"
        }
        $out += "`r`n"

        # ---- rows ----
        $i = 0

        foreach ($row in $dataTable.Rows) {

            if ($script:cancelExport) {
                [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                return
            }

            $out += "|"

            for ($c = 0; $c -lt $dataTable.Columns.Count; $c++) {
                $out += ([string]$row[$c]).PadRight($colWidths[$c]) + "|"
            }

            $out += "`r`n"

            $i++
            $progress.Value = $i

            $percent = [math]::Round(($i / $dataTable.Rows.Count) * 100, 1)
            $eta = Get-ETA $startTime $i $dataTable.Rows.Count

            $lblProgress.Text = "$percent% | ETA: $eta"

            [System.Windows.Forms.Application]::DoEvents()
        }

        Set-Content -Path $dlg.FileName -Value $out

        $progress.Value = $dataTable.Rows.Count
        $lblProgress.Text = "100% complete"

        [System.Windows.Forms.MessageBox]::Show(
            "Export completed successfully!",
            "Done",
            "OK",
            "Information"
        )
    }
})

# ---------------- CSV EXPORT ----------------
$btnExportCsv.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV File (*.csv)|*.csv"

    if ($dlg.ShowDialog() -eq "OK") {

        $script:cancelExport = $false
        $startTime = Get-Date

        $progress.Value = 0
        $progress.Maximum = $dataTable.Rows.Count

        $list = New-Object System.Collections.Generic.List[Object]

        $i = 0
        foreach ($row in $dataTable.Rows) {

            if ($script:cancelExport) {
                [System.Windows.Forms.MessageBox]::Show("Export cancelled.")
                return
            }

            $obj = New-Object PSObject

            for ($c = 0; $c -lt $dataTable.Columns.Count; $c++) {
                $obj | Add-Member -NotePropertyName $dataTable.Columns[$c].ColumnName `
                                  -NotePropertyValue $row[$c]
            }

            $list.Add($obj)

            $i++
            $progress.Value = $i

            $percent = [math]::Round(($i / $dataTable.Rows.Count) * 100, 1)
            $eta = Get-ETA $startTime $i $dataTable.Rows.Count

            $lblProgress.Text = "$percent% | ETA: $eta"

            [System.Windows.Forms.Application]::DoEvents()
        }

        $list | Export-Csv -Path $dlg.FileName -NoTypeInformation

        $progress.Value = $dataTable.Rows.Count
        $lblProgress.Text = "100% complete"

        [System.Windows.Forms.MessageBox]::Show(
            "CSV export completed successfully!",
            "Done",
            "OK",
            "Information"
        )
    }
})

# ---------------- RUN ----------------
[void]$form.ShowDialog()
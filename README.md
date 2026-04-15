# 📊 Profit-Track Data Viewer

A high-performance PowerShell-based tool for parsing, viewing, and exporting large Profit-Track data files without lag or freezing.

Designed specifically to handle **very large files (19,000+ lines)** efficiently while providing a simple GUI for viewing and exporting data.

---

## 🚀 Features

- ⚡ **Optimized for large files** (19,000+ lines and beyond)
- 📂 Load and parse Profit-Track data files quickly
- 📋 Display results in a Windows Forms DataGridView
- 📈 Real-time progress tracking with ETA
- ❌ Cancel long-running operations at any time
- 💾 Export parsed data to:
  - CSV
  - TXT
- 🧠 Memory-efficient streaming (no full file load required)

---

## 📦 Requirements

- Windows
- PowerShell 5.1+ (or PowerShell Core)
- .NET Framework (for Windows Forms)

---

## ▶️ Usage

1. Download or clone the repository:
   ```bash
   git clone https://github.com/yourusername/profit-track-data-viewer.git
   ```

2. Run the script:
   ```powershell
   .\Profit-Track Data Viewer.ps1
   ```

3. Use the GUI to:
   - Open a Profit-Track data file
   - View parsed results
   - Export to CSV or TXT

---

## 📁 File Handling

- Handles very large datasets efficiently
- Processes files line-by-line (streaming)
- Avoids memory spikes and UI freezing

---

## 📤 Exporting

### CSV Export
- Fast and efficient
- Uses direct streaming instead of object pipelines

### TXT Export
- Writes line-by-line using `StreamWriter`
- Avoids large in-memory string builds

---

## ⚙️ How It Works Internally

### 1. File Streaming (Core Performance Feature)
Instead of loading the entire file into memory, the tool uses:

```powershell
System.IO.StreamReader
```

This allows it to:
- Read one line at a time
- Keep memory usage low
- Start processing immediately without waiting for full file load

---

### 2. Parsing Pipeline

Each line is:
1. Read from the file stream
2. Checked against parsing rules
3. Converted into structured data
4. Added to a `DataTable`

To improve performance:
- `BeginLoadData()` is called before bulk inserts
- `EndLoadData()` is called after loading completes

This disables internal change tracking temporarily, making inserts much faster.

---

### 3. UI Optimisation

The GUI uses a Windows Forms `DataGridView`, but with important tweaks:

- ❌ Disabled: `AutoSizeRowsMode = AllCells` *was used in previous builds*
- ❌ Avoids resizing per row
- ✅ Uses batch updates instead of per-line updates

Progress updates are **throttled**:
- UI refresh only happens every *N rows* instead of every row
- Prevents UI thread overload

---

### 4. Progress & ETA Calculation

The script tracks:
- Total file size
- Current read position

From this, it calculates:
- Percentage complete
- Estimated time remaining (ETA)

This is updated periodically to avoid performance hits.

---

### 5. Cancellation System

A cancel flag is checked during processing:

```powershell
if ($cancelRequested) { break }
```

This allows:
- Immediate stop of long operations
- Safe exit without crashing

---

### 6. Export System (High-Speed Writing)

Instead of building large strings in memory:

❌ Slow approach:
```powershell
$text += "line"
```

✅ Optimized approach:
```powershell
System.IO.StreamWriter
```

Benefits:
- Writes directly to disk
- Constant memory usage
- Much faster for large files

---

## 🧯 Troubleshooting

### Script closes immediately or shows errors
- Ensure you're running in **PowerShell**, not CMD
- Run PowerShell as Administrator if needed

### GUI lag on large files
- Make sure you're using the **most recent version**
- Avoid resizing columns manually during load

---

## 📌 Notes

- This tool is designed for **performance first**
- UI responsiveness is prioritised over constant updates
- Large datasets may take time to load, but will not freeze

---

## 📜 License

Creative Commons Attribution-NonCommercial (CC BY-NC 4.0)

---

## 👤 Author

**McSkinnerOG**

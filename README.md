# Auto Printer

A PowerShell script for batch printing PDF, Word, and Excel files silently in Omnissa (VMware) virtual machines. Designed for **Shanghai Jiao Tong University (SJTU) Global College** students.

## Features

- **Silent Printing**: Uses SumatraPDF for completely silent printing (no popup windows)
- **Batch Processing**: Print all files in a folder with a single command
- **Multi-format Support**: PDF, Word (.doc/.docx), Excel (.xls/.xlsx)
- **Progress Tracking**: Shows `[1/10]`, `[2/10]` progress for each file
- **Interruption Handling**: If interrupted, shows which files were not printed
- **Auto Cleanup**: Cleans up the done folder after successful completion

## Prerequisites

### 1. Install SumatraPDF (Required for Silent Printing)

SumatraPDF is a lightweight PDF viewer (~5MB) that supports command-line silent printing.

**Download**: [https://www.sumatrapdfreader.org/download-free-pdf-viewer](https://www.sumatrapdfreader.org/download-free-pdf-viewer)

Install it in your Omnissa Windows VM. The script will automatically detect it.

> **Why SumatraPDF?** Adobe Acrobat will show popup windows during printing, which can be annoying. SumatraPDF's `-print-to-default` flag enables completely silent printing.

### 2. Enable Omnissa Shared Folders

Make sure your Mac folders are mapped to the Windows VM:
- Your Mac Desktop should be accessible as `Z:\Desktop` in the VM
- If not, configure shared folders in Omnissa settings

## Setup

### Step 1: Create Folders on Your Mac

On your **Mac Desktop**, create two folders:

1. **`Print_Queue`** - Put files here that you want to print
2. **`Printed_Done`** - Temporary folder for tracking (auto-cleaned after printing)

### Step 2: Update the Script with Your Student ID

Open `auto-printer.ps1` and find line 15:

```powershell
"C:\Users\<your jaccount ID>\AppData\Local\SumatraPDF\SumatraPDF.exe",  # <-- Replace <your jaccount ID> with your student ID
```

Replace `<your jaccount ID>` with your own student ID (your Windows username in the Omnissa VM).

### Step 3: Place the Script

Put `auto-printer.ps1` in a location accessible from the VM, e.g.:
- `Z:\Documents\myProjects\Auto-printer\auto-printer.ps1`

## Usage

### 1. Add Files to Print Queue

On your **Mac**, drag all the files you want to print into the `Print_Queue` folder on your Desktop.

Supported formats:
- PDF (`.pdf`)
- Word (`.doc`, `.docx`)
- Excel (`.xls`, `.xlsx`)

### 2. Run the Script

In your Omnissa Windows VM, open **PowerShell** and run:

```powershell
powershell -ExecutionPolicy Bypass -File "<file path>"
```

Replace `<file path>` with the full path to `auto-printer.ps1`, e.g.:
- `Z:\Documents\myProjects\Auto-printer\auto-printer.ps1`

### 3. Wait for Completion

The script will:
1. Scan all files in `Print_Queue`
2. Print each file silently to the default printer
3. Move printed files to `Printed_Done`
4. Show a summary when done
5. Clean up `Printed_Done` automatically

Example output:
```
============================================
  Auto Printer - Z:\Desktop\Print_Queue
============================================
[INFO] Default printer: \\printersrv2\JI Printer
[INFO] Found SumatraPDF: C:\Users\<your jaccount ID>\AppData\Local\SumatraPDF\SumatraPDF.exe
----------------------------------------
[1/8] lecture1.pdf
[PRINT] PDF (SumatraPDF silent): Z:\Desktop\Print_Queue\lecture1.pdf
[DONE] Moved to: Z:\Desktop\Printed_Done\lecture1.pdf
----------------------------------------
[2/8] homework.docx
[PRINT] Word document: Z:\Desktop\Print_Queue\homework.docx
[DONE] Moved to: Z:\Desktop\Printed_Done\homework.docx
...
============================================
  ALL DONE!
  Total: 8 | Success: 8 | Failed: 0
============================================
[CLEAN] Cleared Z:\Desktop\Printed_Done (8 file(s) removed)
```

### 4. If Interrupted

If you press `Ctrl+C` or the script crashes, it will show which files were not printed:

```
============================================
  INTERRUPTED! Unprinted files:
============================================
  - lecture5.pdf
  - homework2.docx
============================================
```

You can then re-run the script to print the remaining files.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "No default printer found" | Set your school printer as default in Windows Settings |
| "No PDF reader found" | Install SumatraPDF in the VM |
| Z: drive not found | Enable shared folders in Omnissa settings |
| PDF printing shows popup | Make sure SumatraPDF is installed (Adobe shows popups) |
| Word/Excel printing fails | Install Microsoft Office in the VM |
| Script won't run | Use `-ExecutionPolicy Bypass` flag |

## How It Works

1. **PDF Printing Priority**: SumatraPDF (silent) → Adobe Acrobat → Windows Print verb
2. **Word/Excel Printing**: Uses COM objects to control Office applications silently
3. **File Tracking**: Printed files are moved to `Printed_Done` during processing, then cleaned up on success
4. **Interruption Safety**: `try/finally` ensures unprinted files are always reported

## License

This project is for educational use by SJTU Global College students.

# Auto Printer

A PowerShell script for batch printing PDF, Word, and Excel files silently in Omnissa (VMware) virtual machines. Designed for **Shanghai Jiao Tong University (SJTU) Global College** students.

## Features

- **Silent Printing**: Uses SumatraPDF for completely silent printing (no popup windows)
- **Batch Processing**: Print all files in a folder with a single command
- **Multi-format Support**: PDF, Word (.doc/.docx), Excel (.xls/.xlsx)
- **Progress Bar & Spinner**: Visual progress bar and spinning animation so you always know what's happening
- **Step-by-step Status**: Each file shows `[1/3] Sending... [2/3] Spooling... [3/3] Sent!`
- **Interruption Handling**: If interrupted, shows which files were not printed
- **Auto Cleanup**: Successfully printed files are automatically deleted from the queue
- **One-Click Launch**: Just double-click `run.bat` — no need to type commands

## Prerequisites

### 1. SumatraPDF (Already Included)

SumatraPDF is a lightweight PDF viewer that supports command-line silent printing. The **portable version is already bundled** in the `SumatraPDF/` folder of this project — no download or installation needed.

Because it lives inside the project folder on your shared drive, it survives Omnissa VM resets and you never need to reinstall.

> **Why SumatraPDF?** Adobe Acrobat will show popup windows during printing, which can be annoying. SumatraPDF's `-print-to-default` flag enables completely silent printing.

### 2. Enable Omnissa Shared Folders

Make sure your Mac folders are mapped to the Windows VM:
- Your Mac Desktop should be accessible as `Z:\Desktop` in the VM
- If not, configure shared folders in Omnissa settings

## Setup

### Step 1: Create Folder on Your Mac

On your **Mac Desktop**, create one folder:

- **`Print_Queue`** - Put files here that you want to print

The script will automatically create this folder if it doesn't exist.

### Step 2: Place the Script

Put the project folder in a location accessible from the VM, e.g.:
- `Z:\Documents\myProjects\Auto-printer\`

## Usage

### 1. Add Files to Print Queue

On your **Mac**, drag all the files you want to print into the `Print_Queue` folder on your Desktop.

Supported formats:
- PDF (`.pdf`)
- Word (`.doc`, `.docx`)
- Excel (`.xls`, `.xlsx`)

### 2. Run the Script

#### Easy Way: Double-click `run.bat`

Just double-click `run.bat` in the project folder. That's it!

> **Tip**: Right-click `run.bat` → **Send to** → **Desktop (create shortcut)** to create a desktop shortcut for even faster access.

#### Manual Way: PowerShell

In your Omnissa Windows VM, open **PowerShell** and run:

```powershell
powershell -ExecutionPolicy Bypass -File "<file path>"
```

Replace `<file path>` with the full path to `auto-printer.ps1`, e.g.:
- `Z:\Documents\myProjects\Auto-printer\auto-printer.ps1`

### 3. Watch the Progress

The script will show a progress bar and step-by-step status for each file:

```
============================================
  Auto Printer - Z:\Desktop\Print_Queue
============================================
[INFO] Default printer: \\printersrv2\JI Printer
[INFO] Found SumatraPDF: Z:\Documents\myProjects\Auto-printer\SumatraPDF\SumatraPDF.exe
----------------------------------------
Overall: [#####-------------------------] 16% (1/6)
  File: lecture1.pdf
  [1/3] Sending to SumatraPDF...
  [2/3] Spooling to printer... /
  [2/3] Spooling to printer... Done!
  [3/3] Print job sent!
  -> Deleted (already printed)
----------------------------------------
Overall: [##########-------------------] 33% (2/6)
  File: homework.docx
  [1/3] Opening Word...
  [2/3] Spooling to printer... -
  [2/3] Spooling to printer... Done!
  [3/3] Print job sent!
  -> Deleted (already printed)
...
============================================
  ALL DONE!
Result: [##############################] 100% (6/6)
  Success: 6 | Failed: 0
============================================
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
| "No PDF reader found" | Make sure `SumatraPDF/SumatraPDF.exe` exists in the project folder |
| Z: drive not found | Enable shared folders in Omnissa settings |
| PDF printing shows popup | Make sure portable SumatraPDF is present in `Auto-printer\SumatraPDF\` (Adobe shows popups) |
| Word/Excel printing fails | Install Microsoft Office in the VM |
| Script won't run | Double-click `run.bat` instead, or use `-ExecutionPolicy Bypass` flag |
| Printing seems stuck | Check the spinner animation — if it's still spinning, the print job is being spooled to the printer |

## How It Works

1. **PDF Printing Priority**: Local portable SumatraPDF → Installed SumatraPDF → Adobe Acrobat → Windows Print verb
2. **Word/Excel Printing**: Uses COM objects to control Office applications silently
3. **Visual Feedback**: Progress bar shows overall completion; spinner animation shows when a print job is being spooled
4. **File Management**: Successfully printed files are deleted from the queue; failed files remain for retry
5. **Interruption Safety**: `try/finally` ensures unprinted files are always reported

## Project Structure

```
Auto-printer/
├── auto-printer.ps1     # Main PowerShell script
├── run.bat              # One-click launcher (double-click to run)
├── README.md            # This file
└── SumatraPDF/          # Portable SumatraPDF (survives VM resets)
    └── SumatraPDF.exe
```

## License

This project is for educational use by SJTU Global College students.

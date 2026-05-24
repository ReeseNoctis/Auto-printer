# Auto Printer

A PowerShell script for batch printing PDF, Word, and Excel files silently via a Windows virtual machine. Designed for **Shanghai Jiao Tong University (SJTU) Global College** students.

Works with **macOS**, **Windows**, or **Linux** as your host system — as long as you have a Windows VM with shared folders enabled.

## Features

- **Silent Printing**: Uses SumatraPDF for completely silent printing (no popup windows)
- **Batch Processing**: Print all files in a folder with a single command
- **Multi-format Support**: PDF, Word (.doc/.docx), Excel (.xls/.xlsx)
- **Progress Bar & Spinner**: Visual progress bar and spinning animation so you always know what's happening
- **Step-by-step Status**: Each file shows `[1/3] Sending... [2/3] Spooling... [3/3] Sent!`
- **Interruption Handling**: If interrupted, shows which files were not printed
- **Auto Cleanup**: Successfully printed files are automatically deleted from the queue
- **One-Click Launch**: Just double-click `run.bat` — no need to type commands

## How It Works

```
Your Computer (any OS)              Windows VM
┌─────────────────────┐            ┌──────────────────────┐
│  Desktop/           │            │  Z:\Desktop\         │
│  ├── Print_Queue/   │ ──shared── │  ├── Print_Queue/    │
│  │   ├── doc1.pdf   │   folder   │  │   ├── doc1.pdf    │
│  │   └── doc2.docx  │            │  │   └── doc2.docx   │
│                     │            │                      │
│  Documents/         │            │  Z:\Documents\       │
│  └── Auto-printer/  │ ──shared── │  └── Auto-printer/   │
│      ├── run.bat    │   folder   │      ├── run.bat     │
│      └── ...        │            │      └── ...         │
└─────────────────────┘            └──────────────────────┘
                                         │
                                    Double-click run.bat
                                         │
                                    Files are printed and
                                    deleted from Print_Queue
```

1. You drop files into `Print_Queue` on your **host Desktop**
2. The folder is accessible inside the **Windows VM** via shared folders (e.g. `Z:\Desktop\Print_Queue`)
3. Double-click `run.bat` in the VM — all files print silently and are deleted on success

## Setup (Choose Your Host OS)

All platforms follow the same two steps: (1) clone the repo, (2) create Print_Queue on your Desktop, (3) configure shared folders in your VM software so the Desktop maps to `Z:\Desktop`.

### Step 1: Clone the Project

Open a terminal on your host and clone this repo to a folder that your VM can access (e.g. Documents):

```bash
# macOS / Linux
cd ~/Documents
git clone https://github.com/ReeseNoctis/Auto-printer.git

# Windows (PowerShell)
cd $env:USERPROFILE\Documents
git clone https://github.com/ReeseNoctis/Auto-printer.git
```

### Step 2: Create the Print Queue Folder

Create `Print_Queue` on your Desktop (the script also auto-creates it, so this is optional):

**macOS / Linux:**
```bash
mkdir ~/Desktop/Print_Queue
```

**Windows (PowerShell):**
```powershell
mkdir $env:USERPROFILE\Desktop\Print_Queue
```

### Step 3: Configure Shared Folders

Your **Desktop** and **Documents** folders need to be accessible from inside the Windows VM. The script expects:
- Host `Desktop` → `Z:\Desktop` in the VM
- Host `Documents` → `Z:\Documents` in the VM

Set this up in your VM software:

<details>
<summary><b>macOS + Omnissa (VMware)</b></summary>

1. Open Omnissa settings for your Windows VM
2. Go to **Shared Folders** → enable and add your Mac's `Desktop` and `Documents` folders
3. In the VM, they appear as `Z:\Desktop` and `Z:\Documents`

</details>

<details>
<summary><b>macOS + VMware Fusion</b></summary>

1. Open VMware Fusion → **Virtual Machine** → **Settings** → **Sharing**
2. Enable **Shared Folders** and add your `Desktop` and `Documents` folders
3. In the VM, open File Explorer → **This PC** → they appear under **Network Locations**

If the drive letter isn't `Z:`, you can change it in the VM:
- Open **Disk Management** → right-click the shared folder mapping → **Change Drive Letter and Paths** → assign `Z:`

</details>

<details>
<summary><b>Windows + VMware Workstation</b></summary>

1. Open VMware Workstation → **VM** → **Settings** → **Options** → **Shared Folders**
2. Enable and add your host's `Desktop` and `Documents` folders
3. In the VM, they appear under **Network** in File Explorer
4. To map them to `Z:` drive:
   - Open **Command Prompt** in the VM and run:
     ```
     net use Z: \\vmware-host\Shared Folders\Desktop
     ```
   - Or use **Disk Management** to assign drive letter `Z:`

</details>

<details>
<summary><b>Windows + VirtualBox</b></summary>

1. Open VirtualBox → **Settings** → **Shared Folders**
2. Add your host's `Desktop` folder with name `Desktop`
3. In the Windows VM, open **Command Prompt** and map the drive:
   ```
   net use Z: \\vboxsvr\Desktop
   ```
4. Repeat for your `Documents` folder

</details>

<details>
<summary><b>Linux + VMware Workstation / VirtualBox</b></summary>

Same as Windows host — use the VM software's shared folder feature, then map the drive to `Z:` inside the Windows VM using `net use` or Disk Management.

For VirtualBox:
```
net use Z: \\vboxsvr\Desktop
```

For VMware:
```
net use Z: \\vmware-host\Shared Folders\Desktop
```

</details>

> **Important**: Whatever VM software you use, the host `Desktop` folder must be mapped to `Z:\Desktop` inside the Windows VM. If your setup uses a different drive letter, edit line 9 of `auto-printer.ps1` to match:
> ```powershell
> $watchFolder = "Z:\Desktop\Print_Queue"   # Change Z: to your drive letter
> ```

## Usage

### 1. Add Files to Print Queue

On your **host computer** (any OS), drag files into the `Print_Queue` folder on your Desktop.

Supported formats:
- PDF (`.pdf`)
- Word (`.doc`, `.docx`)
- Excel (`.xls`, `.xlsx`)

### 2. Run the Script

In your **Windows VM**, navigate to the project folder and double-click `run.bat`:

```
Z:\Documents\Auto-printer\run.bat
```

> **Tip**: Right-click `run.bat` → **Send to** → **Desktop (create shortcut)** for one-click access.

#### Manual Way (PowerShell)

In your Windows VM, open **PowerShell** and run:

```powershell
powershell -ExecutionPolicy Bypass -File "Z:\Documents\Auto-printer\auto-printer.ps1"
```

### 3. Watch the Progress

```
============================================
  Auto Printer - Z:\Desktop\Print_Queue
============================================
[INFO] Default printer: \\printersrv2\JI Printer
[INFO] Found SumatraPDF: Z:\Documents\Auto-printer\SumatraPDF\SumatraPDF.exe
----------------------------------------
Overall: [#####-------------------------] 16% (1/6)
  File: lecture1.pdf
  [1/3] Sending to SumatraPDF...
  [2/3] Spooling to printer... /
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

If you press `Ctrl+C` or the script crashes, it shows unprinted files:

```
============================================
  INTERRUPTED! Unprinted files:
============================================
  - lecture5.pdf
  - homework2.docx
============================================
```

Re-run the script to print the remaining files.

## Prerequisites (Inside the Windows VM)

- **SumatraPDF** — bundled in the project, no installation needed
- **Microsoft Office** — required in the VM for Word/Excel printing
- **Default printer** — set your school printer as default in Windows Settings

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "No default printer found" | Set your school printer as default in Windows Settings |
| "No PDF reader found" | Make sure `SumatraPDF/SumatraPDF.exe` exists in the project folder |
| Z: drive not found | Enable shared folders in your VM software settings |
| Shared folder shows wrong drive letter | Use Disk Management in the VM to reassign to `Z:`, or edit the path in `auto-printer.ps1` line 9 |
| PDF printing shows popup | Make sure portable SumatraPDF is present (Adobe shows popups) |
| Word/Excel printing fails | Install Microsoft Office in the VM |
| Script won't run | Double-click `run.bat` instead, or use `-ExecutionPolicy Bypass` flag |

## Project Structure

```
Auto-printer/
├── auto-printer.ps1       # Main PowerShell script
├── run.bat                # One-click launcher (double-click to run)
├── README.md
└── SumatraPDF/            # Portable SumatraPDF (survives VM resets)
    ├── SumatraPDF.exe
    └── SumatraPDF-settings.txt
```

## License

This project is for educational use by SJTU Global College students.

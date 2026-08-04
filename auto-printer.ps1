# ============================================================
# Auto Printer - Batch print PDF/Word/Excel files silently
# Designed for SJTU Global College students using Omnissa VM
# ============================================================

# --- Configuration ---
# The watch folder is mapped from your Mac via Omnissa shared folders (Z: drive)
# Make sure Omnissa shared folders are enabled and your Mac Desktop is mapped to Z:\Desktop
$watchFolder = "Z:\Desktop\Print_Queue"

# --- SumatraPDF paths ---
# Priority: portable (local to project folder) > installed user > installed system
# Portable version survives Omnissa VM resets since it lives on the shared drive
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sumatraPaths = @(
    "$scriptDir\SumatraPDF\SumatraPDF.exe",                               # Portable version in project folder (RECOMMENDED)
    "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe",                        # User install (auto-detected)
    "C:\Program Files\SumatraPDF\SumatraPDF.exe",
    "C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe"
)

$sumatraPath = $null
foreach ($p in $sumatraPaths) {
    if (Test-Path $p) {
        $sumatraPath = $p
        break
    }
}

# --- Performance: copy portable SumatraPDF from network share to local temp ---
# Loading a 20MB EXE from Z:\ over the network on every print adds 10-30s per file.
# Copy once to local %TEMP% at startup — the VM resets on exit anyway, so we
# don't need to clean up.
# Uses robocopy instead of Copy-Item because Copy-Item from a network drive
# (Z:\) to a local path can silently produce nothing in certain VM environments.
if ($sumatraPath -and $sumatraPath.StartsWith($scriptDir)) {
    $localSumatraDir = Join-Path ([System.IO.Path]::GetTempPath()) "SumatraPDF"
    Write-Host "[INFO] Copying SumatraPDF to local temp for faster printing..."

    # Create target directory (robocopy needs it to exist)
    New-Item -ItemType Directory -Path $localSumatraDir -Force -ErrorAction SilentlyContinue | Out-Null

    # Clean any previous stale copy
    Get-ChildItem -Path $localSumatraDir -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

    # robocopy is far more reliable than Copy-Item for network → local copies in VMs.
    # Exit codes 0-7 all indicate success (8+ = failure).
    $sourceDir = "$scriptDir\SumatraPDF"
    cmd /c "robocopy `"$sourceDir`" `"$localSumatraDir`" /E /NFL /NDL /NP /NJH /NJS 2>&1" | Out-Null

    if ($LASTEXITCODE -le 7) {
        # Disable update checks on the local copy to avoid extra network round-trips
        $localSettings = Join-Path $localSumatraDir "SumatraPDF-settings.txt"
        if (Test-Path $localSettings) {
            (Get-Content $localSettings) -replace 'CheckForUpdates = true', 'CheckForUpdates = false' | Set-Content $localSettings
        }

        $localExe = Join-Path $localSumatraDir "SumatraPDF.exe"
        if (Test-Path $localExe) {
            $sumatraPath = $localExe
            Write-Host "[INFO] SumatraPDF ready (local temp): $sumatraPath"
        } else {
            Write-Host "[WARN] robocopy succeeded but EXE not found, using network path"
        }
    }
    else {
        Write-Host "[WARN] robocopy failed (exit code $LASTEXITCODE), using network path"
    }
}

# --- Adobe Acrobat paths (fallback if SumatraPDF is not installed) ---
$adobePaths = @(
    "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
    "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
    "C:\Program Files\Adobe\Reader 11.0\Reader\AcroRd32.exe",
    "C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
    "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
    "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
)

$adobePath = $null
foreach ($p in $adobePaths) {
    if (Test-Path $p) {
        $adobePath = $p
        break
    }
}

# --- Create watch folder if it doesn't exist ---
if (-not (Test-Path $watchFolder)) {
    New-Item -ItemType Directory -Path $watchFolder -Force | Out-Null
    Write-Host "[INFO] Created watch folder: $watchFolder"
}

# --- Detect default printer ---
# Make sure you have set your school printer as the default printer in Windows Settings
$defaultPrinter = (Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }).Name
if ($defaultPrinter) {
    Write-Host "[INFO] Default printer: $defaultPrinter"
} else {
    Write-Host "[WARN] No default printer found! Please set a default printer first."
}

# --- Show which PDF reader was detected ---
# Priority: SumatraPDF (silent) > Adobe Acrobat > Windows native (last resort)
if ($sumatraPath) {
    Write-Host "[INFO] Found SumatraPDF: $sumatraPath"
} elseif ($adobePath) {
    Write-Host "[INFO] Found Adobe: $adobePath"
} else {
    Write-Host "[WARN] No PDF reader found! Please install SumatraPDF or Adobe Acrobat Reader."
}

# --- Progress bar function ---
function Show-ProgressBar {
    param(
        [int]$Current,
        [int]$Total,
        [string]$Label = ""
    )

    $width = 30
    $percent = [math]::Floor(($Current / $Total) * 100)
    $filled = [math]::Floor(($Current / $Total) * $width)
    $empty = $width - $filled

    $bar = "[" + ("#" * $filled) + ("-" * $empty) + "]"
    $status = "$bar $percent% ($Current/$Total)"

    if ($Label) {
        $status = "$Label $status"
    }

    Write-Host $status
}

# --- Spinner function ---
# Shows a spinning animation while waiting for a process to complete.
# Uses [Console]::Write() instead of Write-Host to avoid PowerShell's output
# buffering, which otherwise causes progress to appear in stuttering chunks
# instead of a smooth real-time animation.
function Wait-WithSpinner {
    param(
        [System.Diagnostics.Process]$Process,
        [string]$Message = "Printing",
        [int]$TimeoutSeconds = 120
    )

    $spinner = @("|", "/", "-", "\")
    $i = 0
    $elapsed = 0

    while (!$Process.HasExited -and $elapsed -lt $TimeoutSeconds) {
        $frame = $spinner[$i % 4]
        [Console]::Write("`r  $Message... $frame   ")
        [Console]::Out.Flush()
        Start-Sleep -Milliseconds 200
        $elapsed += 0.2
        $i++
    }

    [Console]::Write("`r  $Message... Done!    `n")
    [Console]::Out.Flush()
}

# --- Print function ---
# Supports: PDF, Word (.doc/.docx), Excel (.xls/.xlsx)
function Print-File {
    param(
        [string]$FilePath
    )

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()

    if ($ext -eq ".pdf") {
        # PDF: SumatraPDF (completely silent) > Adobe (/t flag) > Windows Print verb
        if ($sumatraPath) {
            Write-Host "  [1/2] Sending to SumatraPDF..."
            $proc = Start-Process -FilePath $sumatraPath -ArgumentList "-print-to-default `"$FilePath`"" -PassThru -WindowStyle Hidden
            Write-Host "  [2/2] Spooling to printer..."
            Wait-WithSpinner -Process $proc -Message "  [2/2] Spooling to printer" -TimeoutSeconds 120
            if ($proc -and !$proc.HasExited) {
                Write-Host "  [WARN] SumatraPDF did not exit in time, force-closing..."
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
            Write-Host "  Done!"
        }
        elseif ($adobePath) {
            Write-Host "  [1/2] Sending to Adobe Acrobat..."
            $proc = Start-Process -FilePath $adobePath -ArgumentList "/t `"$FilePath`" `"$defaultPrinter`"" -PassThru -WindowStyle Hidden
            Write-Host "  [2/2] Spooling to printer..."
            Wait-WithSpinner -Process $proc -Message "  [2/2] Spooling to printer" -TimeoutSeconds 120
            if ($proc -and !$proc.HasExited) {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
            Write-Host "  Done!"
        }
        else {
            Write-Host "  [1/2] Sending via Windows Print verb..."
            try {
                $proc = Start-Process -FilePath $FilePath -Verb Print -PassThru -ErrorAction Stop
                Write-Host "  [2/2] Spooling to printer..."
                Wait-WithSpinner -Process $proc -Message "  [2/2] Spooling to printer" -TimeoutSeconds 60
                if ($proc -and !$proc.HasExited) {
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
                Write-Host "  Done!"
            }
            catch {
                Write-Host "  [ERROR] No PDF reader available. Please install SumatraPDF or Adobe Acrobat Reader."
                return $false
            }
        }
    }
    # Word documents: use Word COM object for silent printing (requires Microsoft Office)
    elseif ($ext -in @(".doc", ".docx")) {
        Write-Host "  [1/2] Opening Word..."
        $word = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            Write-Host "  [2/2] Sending to printer..."
            $doc = $word.Documents.Open($FilePath)
            $doc.PrintOut()
            # PrintOut() is synchronous, but we show a brief spinner for feedback
            $spinner = @("|", "/", "-", "\")
            for ($s = 0; $s -lt 15; $s++) {
                [Console]::Write("`r  [2/2] Sending to printer... $($spinner[$s % 4])   ")
                [Console]::Out.Flush()
                Start-Sleep -Milliseconds 200
            }
            [Console]::Write("`r  [2/2] Sending to printer... Done!    `n")
            [Console]::Out.Flush()
            $doc.Close($false)
            $word.Quit()
            Write-Host "  Done!"
        }
        catch {
            Write-Host "  [ERROR] Failed to print Word file: $_"
            if ($word) {
                try { $word.Quit() } catch {}
            }
            return $false
        }
    }
    # Excel documents: use Excel COM object for silent printing (requires Microsoft Office)
    elseif ($ext -in @(".xls", ".xlsx")) {
        Write-Host "  [1/2] Opening Excel..."
        $excel = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            Write-Host "  [2/2] Sending to printer..."
            $wb = $excel.Workbooks.Open($FilePath)
            $ws = $wb.Worksheets.Item(1)
            $ws.PrintOut()
            # PrintOut() is synchronous, but we show a brief spinner for feedback
            $spinner = @("|", "/", "-", "\")
            for ($s = 0; $s -lt 15; $s++) {
                [Console]::Write("`r  [2/2] Sending to printer... $($spinner[$s % 4])   ")
                [Console]::Out.Flush()
                Start-Sleep -Milliseconds 200
            }
            [Console]::Write("`r  [2/2] Sending to printer... Done!    `n")
            [Console]::Out.Flush()
            $wb.Close($false)
            $excel.Quit()
            Write-Host "  Done!"
        }
        catch {
            Write-Host "  [ERROR] Failed to print Excel file: $_"
            if ($excel) {
                try { $excel.Quit() } catch {}
            }
            return $false
        }
    }
    else {
        Write-Host "  [SKIP] Unsupported file type ($ext)"
        return $false
    }

    return $true
}

# --- Main execution ---
Write-Host "============================================"
Write-Host "  Auto Printer - $watchFolder"
Write-Host "============================================"

# Scan all supported files in the queue
$files = Get-ChildItem -Path $watchFolder -File | Where-Object {
    $_.Extension.ToLower() -in @(".pdf", ".doc", ".docx", ".xls", ".xlsx")
}

if ($files.Count -eq 0) {
    Write-Host "[INFO] No files found in Print_Queue. Nothing to print."
    exit 0
}

$totalCount = $files.Count
$successCount = 0
$failCount = 0
$current = 0

# --- Phase 1: Bulk pre-copy all files from Z:\ to local temp ---
# Network reads are slow but sequential. Do them all upfront so the print
# phase runs entirely from local disk — no more per-file "copying..." steps.
$localQueueDir = Join-Path ([System.IO.Path]::GetTempPath()) "Print_Queue_Local"
Write-Host "[INFO] Phase 1/2: Copying $totalCount file(s) to local temp..."
New-Item -ItemType Directory -Path $localQueueDir -Force -ErrorAction SilentlyContinue | Out-Null
Get-ChildItem -Path $localQueueDir -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$localFileMap = @{}  # maps original Z:\path → local C:\...\path
foreach ($file in $files) {
    $localPath = Join-Path $localQueueDir $file.Name
    [Console]::Write("  Copying: $($file.Name)... ")
    [Console]::Out.Flush()
    try {
        [System.IO.File]::Copy($file.FullName, $localPath, $true)
        $localFileMap[$file.FullName] = $localPath
        [Console]::WriteLine("ok")
    } catch {
        [Console]::WriteLine("FAILED — will copy individually later")
    }
    [Console]::Out.Flush()
}

# --- Phase 2: Print from local temp ---
Write-Host ""
Write-Host "[INFO] Phase 2/2: Printing from local temp..."
Write-Host ""

try {
    foreach ($file in $files) {
        $current++
        $filePath = $file.FullName
        $fileName = $file.Name

        # Use pre-copied local file, or copy individually as fallback
        if ($localFileMap.ContainsKey($filePath)) {
            $printPath = $localFileMap[$filePath]
        } else {
            $localPath = Join-Path $localQueueDir $fileName
            Write-Host "  Copying to local temp (fallback)..."
            try {
                [System.IO.File]::Copy($filePath, $localPath, $true)
                $printPath = $localPath
            } catch {
                Write-Host "  [WARN] Copy failed, printing from network: $_"
                $printPath = $filePath
            }
        }

        Write-Host "----------------------------------------"
        Show-ProgressBar -Current $current -Total $totalCount -Label "Overall:"
        Write-Host "  File: $fileName"

        try {
            $success = Print-File -FilePath $printPath

            if ($success) {
                $successCount++
                Remove-Item -Path $filePath -Force
                Write-Host "  -> Deleted from Print_Queue"
            }
            else {
                $failCount++
            }

            # Clean up local temp copy regardless of success/failure
            if ($printPath -ne $filePath) {
                Remove-Item -Path $printPath -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host "  [ERROR] Failed to process ${fileName}: $_"
            $failCount++
        }
    }

    # All files printed successfully
    Write-Host "============================================"
    Write-Host "  ALL DONE!"
    Show-ProgressBar -Current $successCount -Total $totalCount -Label "Result:"
    Write-Host "  Success: $successCount | Failed: $failCount"
    Write-Host "============================================"
}
# If interrupted (Ctrl+C or crash), show which files were not printed
finally {
    $remaining = Get-ChildItem -Path $watchFolder -File | Where-Object {
        $_.Extension.ToLower() -in @(".pdf", ".doc", ".docx", ".xls", ".xlsx")
    }

    if ($remaining.Count -gt 0) {
        Write-Host ""
        Write-Host "============================================"
        Write-Host "  INTERRUPTED! Unprinted files:"
        Write-Host "============================================"
        foreach ($f in $remaining) {
            Write-Host "  - $($f.Name)"
        }
        Write-Host "============================================"
    }

    # Clean up local temp directories
    if ($localQueueDir -and (Test-Path $localQueueDir)) {
        Remove-Item -Recurse -Force $localQueueDir -ErrorAction SilentlyContinue
    }
    if ($localSumatraDir -and (Test-Path $localSumatraDir)) {
        Remove-Item -Recurse -Force $localSumatraDir -ErrorAction SilentlyContinue
    }
}

# ============================================================
# Auto Printer - Batch print PDF/Word/Excel files silently
# Designed for SJTU Global College students using Omnissa VM
# ============================================================

# --- Configuration ---
# The watch folder and done folder are mapped from your Mac via Omnissa shared folders (Z: drive)
# Make sure Omnissa shared folders are enabled and your Mac Desktop is mapped to Z:\Desktop
$watchFolder = "Z:\Desktop\Print_Queue"
$doneFolder = "Z:\Desktop\Printed_Done"

# --- SumatraPDF paths ---
# Replace "<your jaccount ID>" with your own student ID (the Windows username in your Omnissa VM)
$sumatraPaths = @(
    "C:\Users\<your jaccount ID>\AppData\Local\SumatraPDF\SumatraPDF.exe",  # <-- Replace <your jaccount ID> with your student ID
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

# --- Create folders if they don't exist ---
if (-not (Test-Path $watchFolder)) {
    New-Item -ItemType Directory -Path $watchFolder -Force | Out-Null
    Write-Host "[INFO] Created watch folder: $watchFolder"
}

if (-not (Test-Path $doneFolder)) {
    New-Item -ItemType Directory -Path $doneFolder -Force | Out-Null
    Write-Host "[INFO] Created done folder: $doneFolder"
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
            Write-Host "[PRINT] PDF (SumatraPDF silent): $FilePath"
            Start-Process -FilePath $sumatraPath -ArgumentList "-print-to-default `"$FilePath`"" -Wait -WindowStyle Hidden
        }
        elseif ($adobePath) {
            Write-Host "[PRINT] PDF (Adobe Acrobat): $FilePath"
            # Adobe /t flag: print to specified printer and close
            $proc = Start-Process -FilePath $adobePath -ArgumentList "/t `"$FilePath`" `"$defaultPrinter`"" -PassThru -WindowStyle Hidden
            Start-Sleep -Seconds 10
            if ($proc -and !$proc.HasExited) {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            }
        }
        else {
            # Last resort: use Windows shell Print verb (requires a PDF reader to be associated with .pdf)
            Write-Host "[PRINT] PDF (Windows native Print verb): $FilePath"
            try {
                $proc = Start-Process -FilePath $FilePath -Verb Print -PassThru -ErrorAction Stop
                if ($proc) {
                    Start-Sleep -Seconds 8
                    if (!$proc.HasExited) {
                        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                Write-Host "[ERROR] No PDF reader available. Please install SumatraPDF or Adobe Acrobat Reader."
                return $false
            }
        }
    }
    # Word documents: use Word COM object for silent printing (requires Microsoft Office)
    elseif ($ext -in @(".doc", ".docx")) {
        Write-Host "[PRINT] Word document: $FilePath"
        $word = $null
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $doc = $word.Documents.Open($FilePath)
            $doc.PrintOut()
            Start-Sleep -Seconds 5
            $doc.Close($false)
            $word.Quit()
        }
        catch {
            Write-Host "[ERROR] Failed to print Word file: $_"
            if ($word) {
                try { $word.Quit() } catch {}
            }
            return $false
        }
    }
    # Excel documents: use Excel COM object for silent printing (requires Microsoft Office)
    elseif ($ext -in @(".xls", ".xlsx")) {
        Write-Host "[PRINT] Excel document: $FilePath"
        $excel = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $wb = $excel.Workbooks.Open($FilePath)
            $ws = $wb.Worksheets.Item(1)
            $ws.PrintOut()
            Start-Sleep -Seconds 5
            $wb.Close($false)
            $excel.Quit()
        }
        catch {
            Write-Host "[ERROR] Failed to print Excel file: $_"
            if ($excel) {
                try { $excel.Quit() } catch {}
            }
            return $false
        }
    }
    else {
        Write-Host "[SKIP] Unsupported file type ($ext): $FilePath"
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

$printedFiles = @{}

# Print each file, move to done folder on success
try {
    foreach ($file in $files) {
        $current++
        $filePath = $file.FullName
        $fileName = $file.Name

        Write-Host "----------------------------------------"
        Write-Host "[$current/$totalCount] $fileName"

        try {
            $success = Print-File -FilePath $filePath

            if ($success) {
                $successCount++
                $printedFiles[$fileName] = $true
                $destPath = Join-Path $doneFolder $fileName

                # Handle duplicate filenames in done folder
                $counter = 1
                while (Test-Path $destPath) {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                    $ext = [System.IO.Path]::GetExtension($fileName)
                    $destPath = Join-Path $doneFolder "${baseName}_$counter$ext"
                    $counter++
                }

                Move-Item -Path $filePath -Destination $destPath -Force
                Write-Host "[DONE] Moved to: $destPath"
            }
            else {
                $failCount++
            }
        }
        catch {
            Write-Host "[ERROR] Failed to process ${fileName}: $_"
            $failCount++
        }
    }

    # All files printed successfully
    Write-Host "============================================"
    Write-Host "  ALL DONE!"
    Write-Host "  Total: $totalCount | Success: $successCount | Failed: $failCount"
    Write-Host "============================================"

    # Clean up the done folder after successful completion
    if ($successCount -gt 0 -and (Test-Path $doneFolder)) {
        $doneFiles = Get-ChildItem -Path $doneFolder -File
        if ($doneFiles.Count -gt 0) {
            Remove-Item -Path "$doneFolder\*" -Force -Recurse
            Write-Host "[CLEAN] Cleared $doneFolder ($($doneFiles.Count) file(s) removed)"
        }
    }
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
}

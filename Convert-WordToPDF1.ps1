Add-Type -AssemblyName System.Windows.Forms

# --- Prompt for input folder ---
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing Word documents"
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $InputFolder = $folderBrowser.SelectedPath
} else {
    Write-Host "No folder selected. Exiting..."
    exit
}

# --- Prompt for output folder ---
$folderBrowser.Description = "Select the folder to save PDF files"
if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $OutputFolder = $folderBrowser.SelectedPath
} else {
    Write-Host "No output folder selected. Exiting..."
    exit
}

# --- Ensure output folder exists ---
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

Write-Host "Input folder: $InputFolder"
Write-Host "Output folder: $OutputFolder"

# --- Start Word COM object ---
$WordApp = New-Object -ComObject Word.Application
$WordApp.Visible = $false
$WordApp.DisplayAlerts = 0

# --- Get all Word files ---
$WordFiles = Get-ChildItem -Path $InputFolder -Filter *.doc* -File

if ($WordFiles.Count -eq 0) {
    Write-Host "⚠️ No Word documents found in $InputFolder"
    $WordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApp) | Out-Null
    exit
}

Write-Host "Found $($WordFiles.Count) Word files."

# --- Convert each file ---
foreach ($file in $WordFiles) {
    $docPath = $file.FullName
    $pdfPath = Join-Path $OutputFolder ($file.BaseName + ".pdf")

    Write-Host "➡️ Converting: $docPath"
    try {
        $doc = $WordApp.Documents.Open($docPath)
        $doc.SaveAs([ref] ([string]$pdfPath), [ref] 17)  # Cast path to string
        $doc.Close()
        Write-Host "   ✅ Saved: $pdfPath"
    }
    catch {
        Write-Host "   ❌ Failed to convert: $docPath"
        Write-Host "      Error: $_"
    }
}

# --- Quit Word after all files ---
$WordApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WordApp) | Out-Null

Write-Host "All conversions done."

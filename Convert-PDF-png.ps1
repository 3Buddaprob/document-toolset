Add-Type -AssemblyName System.Windows.Forms

# === Prompt for PDF folder ===
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing your PDFs"
$folderBrowser.ShowNewFolderButton = $false
$null = $folderBrowser.ShowDialog()
$InputFolder = $folderBrowser.SelectedPath

if (-not $InputFolder) {
    Write-Host "No folder selected. Exiting."
    exit
}

# === Prompt for Poppler folder ===
$folderBrowser.Description = "Select your Poppler installation folder (root folder)"
$null = $folderBrowser.ShowDialog()
$PopplerRoot = $folderBrowser.SelectedPath

if (-not $PopplerRoot) {
    Write-Host "No Poppler folder selected. Exiting."
    exit
}

# === Find pdftoppm.exe dynamically ===
$PopplerPath = Get-ChildItem -Path $PopplerRoot -Recurse -Filter pdftoppm.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName

if (-not $PopplerPath) {
    Write-Error "Could not find pdftoppm.exe in $PopplerRoot. Please check your Poppler installation."
    exit 1
}

Write-Host "Found Poppler at: $PopplerPath"

# === Output folder ===
$OutputFolder = Join-Path $InputFolder "images"
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# === Convert PDFs ===
Get-ChildItem -Path $InputFolder -Filter *.pdf | ForEach-Object {
    $pdf = $_.FullName
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    $outputBase = Join-Path $OutputFolder $basename

    Write-Host "Converting $($_.Name)..."
    & $PopplerPath -png -r 300 $pdf $outputBase
    Write-Host "Done: $($_.Name)"
}

Write-Host "All PDFs have been converted!"

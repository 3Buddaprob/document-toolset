# Load Windows Forms for folder/file selection
Add-Type -AssemblyName System.Windows.Forms

# Default values
$DefaultPhotoFolder = "\\gh2vmi_VNAS07\MyDocs$\hyde.chiu\Downloads\drive-download-20250905T112134Z-1-001"
$DefaultExifToolPath = "\\GH2VMI_VNAS07\desktop$\hyde.chiu\Desktop\New folder\exiftool-13.34_64\exiftool.exe"
$DefaultAPIKey = "pk.fa6ac10d5ebb9ecf62c38ae2ed03309d"

# Ask user to select the folder with photos
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing your photos"
$folderBrowser.ShowNewFolderButton = $false

if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $PhotoFolder = $folderBrowser.SelectedPath
} else {
    Write-Host "No folder selected, using default: $DefaultPhotoFolder"
    $PhotoFolder = $DefaultPhotoFolder
}

# Ask user to select the ExifTool executable
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "exe files (*.exe)|*.exe"
$openFileDialog.Title = "Select the ExifTool executable"

if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $ExifToolPath = $openFileDialog.FileName
} else {
    Write-Host "No ExifTool executable selected, using default: $DefaultExifToolPath"
    $ExifToolPath = $DefaultExifToolPath
}

# Ask user to input their LocationIQ API key
$APIKey = Read-Host -Prompt "Enter your LocationIQ API Key (press Enter to use default)"
if ([string]::IsNullOrWhiteSpace($APIKey)) {
    Write-Host "No API key entered, using default."
    $APIKey = $DefaultAPIKey
}


function Convert-DMSToDecimal($dms) {
    # Use double quotes outside, escape " with backtick
    if ($dms -match "(\d+) deg (\d+)' ([\d\.]+)`" ([NSEW])") {
        $deg = [double]$matches[1]
        $min = [double]$matches[2]
        $sec = [double]$matches[3]
        $dir = $matches[4]

        $decimal = $deg + ($min/60) + ($sec/3600)
        if ($dir -eq 'S' -or $dir -eq 'W') { $decimal = -$decimal }
        return $decimal
    } else {
        Write-Host "Failed to parse DMS: $dms"
        return $null
    }
}

# Get all JPG files
$photos = Get-ChildItem -Path $PhotoFolder -Filter "*.jpg"
Write-Host "Found $($photos.Count) photos.`n"

# Counter for unknown locations
$unknownCounter = 1

foreach ($photo in $photos) {
    Write-Host "Processing photo: $($photo.Name)"

    # Get GPS coordinates
    $gpsOutput = & $ExifToolPath -s -s -s -GPSLatitude -GPSLongitude $photo.FullName
    Write-Host "Raw ExifTool output:`n$gpsOutput"

    # Split latitude / longitude
    $coords = $gpsOutput -split '\s(?=\d+ deg)'
    if ($coords.Count -eq 2) {
        $latDMS = $coords[0].Trim()
        $lonDMS = $coords[1].Trim()
        Write-Host "Extracted Latitude DMS: $latDMS"
        Write-Host "Extracted Longitude DMS: $lonDMS"
    } else {
        $latDMS = $null
        $lonDMS = $null
        Write-Host "Failed to split latitude and longitude from: $gpsOutput"
    }

    if (-not $latDMS -or -not $lonDMS) {
        $roadName = "UnknownLocation $unknownCounter"
        $unknownCounter++
        Write-Host "No GPS info found, assigned road name: $roadName"
    } else {
        # Convert to decimal
        $latDec = Convert-DMSToDecimal $latDMS
        $lonDec = Convert-DMSToDecimal $lonDMS
        Write-Host "Decimal Latitude: $latDec"
        Write-Host "Decimal Longitude: $lonDec"

        # Call LocationIQ
        try {
            $url = "https://us1.locationiq.com/v1/reverse.php?key=$APIKey&lat=$latDec&lon=$lonDec&format=json"
            Write-Host "Calling LocationIQ API: $url"
            $response = Invoke-RestMethod -Uri $url -Method Get
            if ($response.address.road) {
                $roadName = $response.address.road
            } else {
                $roadName = "UnknownLocation $unknownCounter"
                $unknownCounter++
            }
            Write-Host "LocationIQ returned road name: $roadName"
        } catch {
            $roadName = "UnknownLocation $unknownCounter"
            $unknownCounter++
            Write-Host "Error calling LocationIQ, assigned road name: $roadName"
        }
    }

    # Sanitize file name
    $safeName = ($roadName -replace '[\\\/:\*\?"<>\|]', '') + ".jpg"
    $newName = Join-Path -Path $PhotoFolder -ChildPath $safeName

    # Handle duplicates
    $i = 1
    while (Test-Path $newName) {
        $safeName = ($roadName -replace '[\\\/:\*\?"<>\|]', '') + " $i.jpg"
        $newName = Join-Path -Path $PhotoFolder -ChildPath $safeName
        $i++
    }

    # Rename
    Rename-Item -Path $photo.FullName -NewName $newName
    Write-Host "Renamed $($photo.Name) -> $safeName`n"
}

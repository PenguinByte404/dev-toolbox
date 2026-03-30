[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$FilePath
)

# Prompt for the file path if it wasn't provided in the command line
if ([string]::IsNullOrWhiteSpace($FilePath)) {
    $FilePath = Read-Host "Please enter the full path to your .xlsx or .xlsm file"
    # Remove any accidental quotation marks the user might have pasted
    $FilePath = $FilePath.Replace('"', '')
}

# Ensure the file exists before proceeding
if (-not (Test-Path $FilePath)) {
    Write-Host "Error: Cannot find file at path '$FilePath'" -ForegroundColor Red
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Make a backup of the original file
$BackupPath = "$FilePath.backup"
Copy-Item -Path $FilePath -Destination $BackupPath -Force
Write-Host "Backup created at: $BackupPath" -ForegroundColor Cyan

# Load the .NET compression assembly
Add-Type -AssemblyName System.IO.Compression.FileSystem

try {
    # Open the Excel file directly as a Zip Archive
    $zipArchive = [System.IO.Compression.ZipFile]::Open($FilePath, [System.IO.Compression.ZipArchiveMode]::Update)
    
    # Filter for all worksheet XML files
    $worksheets = $zipArchive.Entries | Where-Object { $_.FullName -match "^xl/worksheets/sheet.*\.xml$" }
    
    $protectionRemoved = $false

    foreach ($sheet in $worksheets) {
        $stream = $sheet.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()
        
        # Check if the sheetProtection tag exists
        if ($content -match '<sheetProtection[^>]*/>') {
            Write-Host "Protection found in $($sheet.Name). Removing..." -ForegroundColor Yellow
            
            # Strip out the entire sheetProtection tag
            $newContent = $content -replace '<sheetProtection[^>]*/>', ''
            
            # Rewrite the file inside the archive
            $stream.Position = 0
            $stream.SetLength(0)
            $writer = New-Object System.IO.StreamWriter($stream)
            $writer.Write($newContent)
            $writer.Flush()
            
            $protectionRemoved = $true
        }
        $stream.Close()
    }

    if ($protectionRemoved) {
        Write-Host "Success! Sheet protection was removed." -ForegroundColor Green
    } else {
        Write-Host "No sheet protection tags were found in this workbook." -ForegroundColor Gray
    }

} catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
} finally {
    if ($zipArchive) {
        $zipArchive.Dispose()
    }
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

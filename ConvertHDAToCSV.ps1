# HDA to CSV Converter Script - Batch Processing
# Converts multiple Oracle UCM HDA export files to a single CSV format

param(
    [Parameter(Mandatory=$true)]
    [string]$FolderPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFileName = "combined_export.csv"
)

function Convert-HDAToCSV {
    param(
        [string]$InputPath,
        [bool]$IncludeHeader = $true
    )
    
    try {
        # Read the HDA file
        $content = Get-Content -Path $InputPath -Raw -Encoding UTF8
        
        # Extract the ResultSet section
        $resultSetPattern = '@ResultSet ExportResults\s*\r?\n(.*?)@end'
        $resultSetMatch = [regex]::Match($content, $resultSetPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        if (-not $resultSetMatch.Success) {
            throw "Could not find ResultSet section in HDA file: $InputPath"
        }
        
        $resultSetContent = $resultSetMatch.Groups[1].Value
        
        # Split into lines - keep ALL lines including empty ones
        $allLines = $resultSetContent -split "`r?`n"
        
        # First line should be the number of columns
        $numColumns = [int]$allLines[0].Trim()
        
        # Extract column definitions
        $columns = @()
        
        for ($i = 1; $i -le $numColumns; $i++) {
            $columnLine = $allLines[$i].Trim()
            $parts = $columnLine -split '\s+'
            $columnName = $parts[0]
            $columns += $columnName
        }
        
        # Find where actual data starts
        $dataStartIndex = $numColumns + 1
        
        # Get all data lines (including empty ones that represent empty fields)
        # Remove only the final @end line if it exists
        $endIndex = $allLines.Count - 1
        if ($allLines[$endIndex].Trim() -eq "@end" -or $allLines[$endIndex].Trim() -eq "") {
            $endIndex = $allLines.Count - 2
        }
        
        $dataLines = $allLines[$dataStartIndex..$endIndex]
        
        # Process data - each record should have exactly $numColumns fields
        $records = @()
        $currentRecord = @()
        $lineIndex = 0
        
        while ($lineIndex -lt $dataLines.Count) {
            # Get the field value (could be empty)
            $fieldValue = $dataLines[$lineIndex]
            $currentRecord += $fieldValue
            
            # Check if we have a complete record
            if ($currentRecord.Count -eq $numColumns) {
                $records += ,$currentRecord  # The comma forces it to be treated as a single array element
                $currentRecord = @()
            }
            
            $lineIndex++
        }
        
        # Handle any incomplete record at the end
        if ($currentRecord.Count -gt 0) {
            Write-Warning "Last record incomplete with $($currentRecord.Count) fields. Expected $numColumns fields."
            
            # Pad with empty values
            while ($currentRecord.Count -lt $numColumns) {
                $currentRecord += ""
            }
            $records += ,$currentRecord
        }
        
        # Create CSV content
        $csvLines = @()
        
        # Add header only if requested (first file)
        if ($IncludeHeader) {
            $headerLine = ($columns | ForEach-Object { 
                if ($_ -match '[,"\r\n]') { 
                    '"' + ($_ -replace '"', '""') + '"' 
                } else { 
                    $_ 
                } 
            }) -join ','
            $csvLines += $headerLine
        }
        
        # Add data rows
        $rowNumber = 1
        foreach ($record in $records) {
            if ($record.Count -ne $numColumns) {
                Write-Warning "Record $rowNumber has $($record.Count) fields instead of $numColumns in file: $InputPath"
            }
            
            $csvRow = ($record | ForEach-Object {
                $field = if ($_ -eq $null) { "" } else { $_.ToString().Trim() }
                $field = $field -replace '[\r\n]', ' '  # Replace line breaks with spaces
                if ($field -match '[,"\r\n]' -or $field -match '^\s|\s$') {
                    '"' + ($field -replace '"', '""') + '"'
                } else {
                    $field
                }
            }) -join ','
            $csvLines += $csvRow
            $rowNumber++
        }
        
        return @{
            CsvLines = $csvLines
            RecordCount = $records.Count
            ColumnCount = $numColumns
        }
    }
    catch {
        Write-Error "Error converting HDA file '$InputPath': $($_.Exception.Message)"
        return $null
    }
}

# Main execution
if ($FolderPath -eq "") {
    Write-Host "Usage: .\Convert-HDAToCSV.ps1 -FolderPath 'path\to\folder' [-OutputFileName 'output.csv']"
    exit 1
}

if (-not (Test-Path $FolderPath -PathType Container)) {
    Write-Error "Folder not found: $FolderPath"
    exit 1
}

# Find all HDA files in the folder, excluding docmetadefinition.hda
$hdaFiles = Get-ChildItem -Path $FolderPath -Filter "*.hda" | Where-Object { $_.Name -ne "docmetadefinition.hda" } | Sort-Object Name

if ($hdaFiles.Count -eq 0) {
    Write-Error "No HDA files found in folder (excluding docmetadefinition.hda): $FolderPath"
    exit 1
}

Write-Host "Starting HDA to CSV batch conversion..."
Write-Host "Input folder: $FolderPath"
Write-Host "Found $($hdaFiles.Count) HDA files (excluding docmetadefinition.hda)"

# Combine all CSV content
$allCsvLines = @()
$totalRecords = 0
$totalColumns = 0
$firstFile = $true

foreach ($hdaFile in $hdaFiles) {
    Write-Host "Processing: $($hdaFile.Name)"
    
    $result = Convert-HDAToCSV -InputPath $hdaFile.FullName -IncludeHeader $firstFile
    
    if ($result -eq $null) {
        Write-Error "Failed to process file: $($hdaFile.Name)"
        continue
    }
    
    $allCsvLines += $result.CsvLines
    $totalRecords += $result.RecordCount
    
    if ($firstFile) {
        $totalColumns = $result.ColumnCount
        $firstFile = $false
    }
    
    Write-Host "  - Processed $($result.RecordCount) records"
}

if ($allCsvLines.Count -eq 0) {
    Write-Error "No data was processed successfully."
    exit 1
}

# Write combined output
$outputPath = Join-Path -Path $FolderPath -ChildPath $OutputFileName
$allCsvLines | Out-File -FilePath $outputPath -Encoding UTF8 -Force

Write-Host "`nBatch conversion completed successfully!"
Write-Host "Input folder: $FolderPath"
Write-Host "Output file: $outputPath"
Write-Host "Files processed: $($hdaFiles.Count)"
Write-Host "Total columns: $totalColumns"
Write-Host "Total records: $totalRecords"
Write-Host "`nCombined CSV saved to: $outputPath"
Add-Type -AssemblyName System.Data.OracleClient
Add-Type -AssemblyName System.Drawing

# ================================
# CONFIGURATION
# ================================
$username = "[USER]"
$password = "[PASSWORD]"
$data_source = "[HOST]:[PORT]/[SERVICE_NAME]"

# Filtering Query
$query = "SELECT F.BFILEDATA, R.DID, R.DWEBEXTENSION FROM DOC_OCS.REVISIONS R JOIN DOC_OCS.FILESTORAGE F ON F.DID = R.DID WHERE (R.DINDATE < TO_DATE('[START_DATE]', 'YYYY-MM-DD') AND R.DINDATE >= TO_DATE('[END_DATE]', 'YYYY-MM-DD')) AND R.DWEBEXTENSION IN ('tif','tiff')"

# Output Directories
$output_dir = "[EXPORT_PATH}" #C:\Export\
$logFile = Join-Path $output_dir "PageCount_Report.log"

# Initialize Counters
$grandTotalPages = 0
$totalDocuments = 0

# ================================
# CONNECT TO ORACLE
# ================================
$connection_string = "User Id=$username;Password=$password;Data Source=$data_source"

try {
    $con = New-Object System.Data.OracleClient.OracleConnection($connection_string)
    $con.Open()

    # ================================
    # EXECUTE QUERY
    # ================================
    $cmd = $con.CreateCommand()
    $cmd.CommandText = $query
    $reader = $cmd.ExecuteReader()
    
    # Initialize Log with Header
    "dID | PageCount | Status" | Out-File $logFile -Encoding UTF8

    # ================================
    # PROCESS EACH ROW
    # ================================
    while ($reader.Read()) {
        $id = $reader["DID"]
        $blob = $reader["BFILEDATA"]
        
        # Verify that $blob is a valid System.Byte[]
        if ($blob -is [System.Byte[]] -and $blob.Length -gt 0) {
            
            # Create a MemoryStream from the byte array (Processing in RAM)
            $memStream = New-Object System.IO.MemoryStream(,$blob)
            $image = $null
            
            try {
                # Load the image from memory
                $image = [System.Drawing.Image]::FromStream($memStream)
                
                # Count the pages (frames)
                $pageDimension = [System.Drawing.Imaging.FrameDimension]::Page
                $pageCount = $image.GetFrameCount($pageDimension)
                
                # Increment Totals
                $grandTotalPages += $pageCount
                $totalDocuments++

                # Log Success
                $logEntry = "{0} | {1} | Success" -f $id, $pageCount
                Write-Host $logEntry -ForegroundColor Green
                Add-Content -Path $logFile -Value $logEntry

            } catch {
                # Log Image Processing Error (e.g., corrupt TIFF)
                $logEntry = "{0} | 0 | Error: {1}" -f $id, $_.Exception.Message
                Write-Warning $logEntry
                Add-Content -Path $logFile -Value $logEntry
            } finally {
                # CLEANUP MEMORY IMMEDIATELY
                if ($image) { $image.Dispose() }
                if ($memStream) { $memStream.Dispose() }
            }

        } else {
            # Log Empty Blob
            $logEntry = "{0} | 0 | Error: Empty BLOB" -f $id
            Write-Warning $logEntry
            Add-Content -Path $logFile -Value $logEntry
        }
    }

    $reader.Close()
    
    # ================================
    # FINAL CALCULATIONS & REPORT
    # ================================
    
    # Avoid Division by Zero Error
    if ($totalDocuments -gt 0) {
        # Calculate Average (Rounded to 2 decimal places)
        $averagePages = [math]::Round(($grandTotalPages / $totalDocuments), 2)
    } else {
        $averagePages = 0
    }

    $footer = @"
------------------------------------------------
FINAL REPORT SUMMARY
------------------------------------------------
TOTAL DOCUMENTS PROCESSED: $totalDocuments
GRAND TOTAL PAGES        : $grandTotalPages
AVERAGE PAGES PER DOC    : $averagePages
------------------------------------------------
"@
    
    # Write footer to log and console
    Add-Content -Path $logFile -Value $footer
    Write-Host $footer -ForegroundColor Cyan

# ================================
# ERROR HANDLING
# ================================
} catch {
    Write-Error ("Database Exception: {0}`n{1}" -f $con.ConnectionString, $_.Exception.ToString())
    $logEntry = "{0} | Database Exception: {1}`n{2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $con.ConnectionString, $_.Exception.ToString()
    Add-Content -Path $logFile -Value $logEntry

# ================================
# CLEANUP
# ================================
} finally {
    if ($con.State -eq 'Open') { $con.Close() }
}
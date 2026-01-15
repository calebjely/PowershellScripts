$libPath = $PSScriptRoot
Add-Type -AssemblyName System.Data.OracleClient
Add-Type -AssemblyName System.Drawing
Add-Type -Path (Join-Path $libPath "itext.io.dll")
Add-Type -Path (Join-Path $libPath "itext.kernel.dll")
Add-Type -Path (Join-Path $libPath "itext.layout.dll")

# ============================================
# CONFIGURATION
# ============================================

# Connection Details
$username = "[USER]"
$password = "[PASSWORD]"
$data_source = "[HOST]:[PORT]/[SERVICE_NAME]"

# Query
$query = "SELECT F.BFILEDATA, R.DID, R.DWEBEXTENSION FROM DOC_OCS.REVISIONS R JOIN DOC_OCS.FILESTORAGE F ON F.DID = R.DID WHERE R.DINDATE > TO_DATE('[END_DATE]', 'YYYY-MM-DD') AND (R.DINDATE <= TO_DATE('[START_DATE]', 'YYYY-MM-DD') AND R.DWEBEXTENSION IN ('tif','tiff','pdf')"

# Output Log Location
$logFile = Join-Path $PSScriptRoot "PageCount_Report.log"

# ============================================
# EXECUTION
# ============================================

# Initialize Counters
$totalDocs = 0
$grandTotalPages = 0
$runningTotal = 0

try {
    $con = New-Object System.Data.OracleClient.OracleConnection("User Id=$username;Password=$password;Data Source=$data_source")
    $con.Open()

    $cmd = $con.CreateCommand()
    $cmd.CommandText = $query
    $reader = $cmd.ExecuteReader()
    
    # Header for the log file
    "dID | Type | PagesInDoc | RunningTotal | Status" | Out-File $logFile -Encoding UTF8

    while ($reader.Read()) {
        $id = $reader["DID"]
        $blob = $reader["BFILEDATA"]
        $ext = $reader["DWEBEXTENSION"].ToString().ToLower().Trim()
        
        if ($blob -is [System.Byte[]] -and $blob.Length -gt 0) {
            $memStream = New-Object System.IO.MemoryStream(,$blob)
            $pageCount = 0
            
            try {
                if ($ext -eq "pdf") {
                    # iText 7 Logic
                    $pdfReader = New-Object itext.kernel.pdf.PdfReader($memStream)
                    $pdfDoc = New-Object itext.kernel.pdf.PdfDocument($pdfReader)
                    $pageCount = $pdfDoc.GetNumberOfPages()
                    
                    $pdfDoc.Close()
                    $pdfReader.Close()
                } 
                elseif ($ext -in "tif", "tiff") {
                    # System.Drawing Logic
                    $image = [System.Drawing.Image]::FromStream($memStream)
                    $pageCount = $image.GetFrameCount([System.Drawing.Imaging.FrameDimension]::Page)
                    $image.Dispose()
                }

                # Update running statistics
                $totalDocs++
                $grandTotalPages += $pageCount
                $runningTotal += $pageCount

                # Log to console and file
                $logEntry = "{0} | {1} | {2} | {3} | Success" -f $id, $ext, $pageCount, $runningTotal
                Write-Host $logEntry -ForegroundColor Green
                Add-Content -Path $logFile -Value $logEntry

            } catch {
                $logEntry = "{0} | {1} | 0 | {2} | Error: {3}" -f $id, $ext, $runningTotal, $_.Exception.Message
                Write-Warning $logEntry
                Add-Content -Path $logFile -Value $logEntry
            } finally {
                $memStream.Dispose()
            }
        }
    }
    
    # Final Summary Calculation
    $avg = if ($totalDocs -gt 0) { [math]::Round(($grandTotalPages / $totalDocs), 2) } else { 0 }
    $footer = @"

------------------------------------------------
FINAL SUMMARY
------------------------------------------------
TOTAL DOCUMENTS      : $totalDocs
GRAND TOTAL PAGES    : $grandTotalPages
AVERAGE PAGES PER DOC: $avg
------------------------------------------------
"@
    Add-Content -Path $logFile -Value $footer
    Write-Host $footer -ForegroundColor Cyan

} catch {
    Write-Error "Database Error: $($_.Exception.Message)"
} finally {
    if ($con -and $con.State -eq 'Open') { $con.Close() }
}
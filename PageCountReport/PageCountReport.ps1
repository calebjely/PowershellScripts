# ============================================
# 1. LIBRARIES
# ============================================
$files = @("BouncyCastle.Crypto.dll", "itextsharp.dll")

foreach ($file in $files) {
    $fullPath = Join-Path $PSScriptRoot $file
    if (Test-Path $fullPath) {
        # Using LoadFrom is safer for iterative testing in the same session
        [System.Reflection.Assembly]::LoadFrom($fullPath) | Out-Null
    }
}

Add-Type -AssemblyName System.Data.OracleClient
Add-Type -AssemblyName System.Drawing

# ============================================
# 2. CONFIGURATION VARIBLES
# ============================================

$username = "[USER]"
$password = "[PASSWORD]"
$data_source = "[HOST]:[PORT]/[SERVICE_NAME]"
$searchFilters = "REVISIONS.DINDATE >= TO_DATE('[OLD_DATE]', 'YYYY-MM-DD') AND REVISIONS.DINDATE <= TO_DATE('[NEW_DATE]', 'YYYY-MM-DD') AND REVISIONS.DWEBEXTENSION IN ('tif','tiff','pdf')"

# ============================================
# 3. STATIC VARIABLES
# ============================================

$timestamp = Get-Date -Format "yyyy.MM.dd-HH.mm.ss"
$logFile = Join-Path $PSScriptRoot "$timestamp-PageCount.log"
$countQuery = "SELECT COUNT(*) FROM REVISIONS WHERE $searchFilters"
$dataQuery = "SELECT FILESTORAGE.BFILEDATA, REVISIONS.DID, REVISIONS.DWEBEXTENSION FROM REVISIONS JOIN FILESTORAGE ON FILESTORAGE.DID = REVISIONS.DID WHERE $searchFilters"

# ============================================
# 4. EXECUTION
# ============================================

try {
    $con = New-Object System.Data.OracleClient.OracleConnection("User Id=$username;Password=$password;Data Source=$data_source")
    $con.Open()

    $cmdCount = $con.CreateCommand()
    $cmdCount.CommandText = $countQuery
    $totalDocsFound = $cmdCount.ExecuteScalar()
	$header = @"
	
Total Documents Found : $totalDocsFound

-----------------------------------------------------
Progress | dID | Type | Pages | RunningTotal | Status
-----------------------------------------------------
"@

    $header | Out-File $logFile -Encoding UTF8
	Write-Host $header -ForegroundColor Cyan

    $cmdData = $con.CreateCommand()
    $cmdData.CommandText = $dataQuery
    $reader = $cmdData.ExecuteReader() 

    $processedDocs = 0; $grandTotalPages = 0; $runningTotal = 0

    while ($reader.Read()) {
        $processedDocs++
        $id = $reader["DID"]
        $blob = $reader["BFILEDATA"]
        $ext = $reader["DWEBEXTENSION"].ToString().ToLower().Trim()
        
        if ($blob -is [System.Byte[]] -and $blob.Length -gt 0) {
            $memStream = New-Object System.IO.MemoryStream(,$blob)
            $pageCount = 0
            
            try {
                if ($ext -eq "pdf") {
                    $pdfReader = New-Object iTextSharp.text.pdf.PdfReader($memStream)
                    $pageCount = $pdfReader.NumberOfPages
                    $pdfReader.Close()
                } 
                elseif ($ext -in "tif", "tiff") {
                    $image = [System.Drawing.Image]::FromStream($memStream)
                    $pageCount = $image.GetFrameCount([System.Drawing.Imaging.FrameDimension]::Page)
                    $image.Dispose()
                }

                $grandTotalPages += $pageCount
                $runningTotal += $pageCount
                $entry = "{0} of {1} | {2} | {3} | {4} | {5} | Success" -f $processedDocs, $totalDocsFound, $id, $ext, $pageCount, $runningTotal
                Write-Host $entry -ForegroundColor Green
                $entry | Add-Content $logFile
            } catch {
                $entry = "{0} of {1} | {2} | {3} | 0 | {4} | Error: $($_.Exception.Message)" -f $processedDocs, $totalDocsFound, $id, $ext, $runningTotal
                Write-Warning $entry
                $entry | Add-Content $logFile
            } finally {
                if ($memStream) { $memStream.Dispose() }
            }
        }
    }
    $reader.Close()
    
    # Calculate Average
    $avgPages = 0
    if ($processedDocs -gt 0) {
        $avgPages = [math]::Round(($grandTotalPages / $processedDocs), 2)
    }

    # Final Summary with HR
    $footer = @"

--------------------------------------------------
FINAL SUMMARY
--------------------------------------------------
Total Documents: $processedDocs
Total Pages:     $grandTotalPages
Avg Pages/Doc:   $avgPages
--------------------------------------------------
"@

    Add-Content $logFile $footer
    Write-Host $footer -ForegroundColor Cyan

} finally {
    if ($con -and $con.State -eq 'Open') { $con.Close() }
}
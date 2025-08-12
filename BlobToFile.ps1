Add-Type -AssemblyName System.Data.OracleClient

# ================================
# CONFIGURATION
# ================================
$username = "[USER]"
$password = "[PASSWORD]"
$data_source = "[HOST]:[PORT]/[SERVICE_NAME]"
$query = "SELECT F.BFILEDATA, R.DID, R.DWEBEXTENSION FROM [USER].REVISIONS R JOIN [USER].FILESTORAGE F ON F.DID = R.DID WHERE R.DINDATE < TO_DATE('[DATE]', 'YYYY-MM-DD') OR R.DINDATE >= TO_DATE('[DATE]', 'YYYY-MM-DD')"
$output_dir = "D:\Export\"
$logFile = Join-Path $output_dir "Export.log"

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

    # ================================
    # PROCESS EACH ROW
    # ================================
    while ($reader.Read()) {
        $id = $reader["DID"]
        $extension = $reader["DWEBEXTENSION"]
		$extension = (($extension.Trim().ToLower()) -replace '[^a-z0-9]', '')
        $blob = $reader["BFILEDATA"]
		
		# Verify that $blob is a valid System.Byte[]
		if ($blob -is [System.Byte[]] -and $blob.Length -gt 0) {
			$outputPath = Join-Path $output_dir "$id.$extension"
			[System.IO.File]::WriteAllBytes($outputPath, $blob)
			
			# Log the Export of The File
            Write-Host "$outputPath has been Exported."
			$logEntry = "{0} | {1} has been Exported." -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $outputPath
            Add-Content -Path $logFile -Value $logEntry
		} else {
		
			# Log the Export Error
			Write-Warning "dID $id is not a valid BLOB or is empty."
			$logEntry = "{0} | dID {1} is not a valid BLOB or is empty.." -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $outputPath
            Add-Content -Path $logFile -Value $logEntry
		}
    }

    $reader.Close()

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
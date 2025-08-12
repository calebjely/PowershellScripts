Add-Type -AssemblyName System.Data.OracleClient

$username = "[USER]"
$password = "[PASSWORD]"
$data_source = "[HOST]:[PORT]/[SERVICE_NAME]"

$connection_string = "User Id=$username;Password=$password;Data Source=$data_source"

try {
    $con = New-Object System.Data.OracleClient.OracleConnection($connection_string)
    $con.Open()

    # Query BILLING_1232
    $cmd1 = $con.CreateCommand()
    $cmd1.CommandText = "SELECT AGENCY, USERCOUNT FROM OTDS.BILLING_1232"
    $reader1 = $cmd1.ExecuteReader()
    $dt1232 = New-Object System.Data.DataTable
    $dt1232.Load($reader1)
    $dt1232.Columns.Add("CHARGE CODE", [string])
    foreach ($row in $dt1232.Rows) {
        $row["CHARGE CODE"] = "1232"
        $row["USERCOUNT"] = [int]$row["USERCOUNT"]
    }

    # Query BILLING_1232C
    $cmd2 = $con.CreateCommand()
    $cmd2.CommandText = "SELECT AGENCY, USERCOUNT FROM OTDS.BILLING_1232C"
    $reader2 = $cmd2.ExecuteReader()
    $dt1232C = New-Object System.Data.DataTable
    $dt1232C.Load($reader2)
    $dt1232C.Columns.Add("CHARGE CODE", [string])
    foreach ($row in $dt1232C.Rows) {
        $row["CHARGE CODE"] = "1232C"
        $row["USERCOUNT"] = [int]$row["USERCOUNT"]
    }
    
	# Combine both tables
	$combined = $dt1232.Copy()
	$combined.Merge($dt1232C)

	# Add CHARGE CODE as first column in new export table
	$exportTable = New-Object System.Data.DataTable
	$exportTable.Columns.Add("CHARGE CODE", [string])
	$exportTable.Columns.Add("BILLCODE", [string])
	$exportTable.Columns.Add("COUNT", [int])

	# Dictionary for 1232A aggregation
	$dict1232A = @{}

	# Step 1: Transform originals (set COUNT to 1, prepare 1232A data)
	foreach ($row in $combined.Rows) {
		$agency = $row["AGENCY"]
		$originalCode = $row["CHARGE CODE"]
		$originalCount = [int]$row["USERCOUNT"]
		$adjusted = $originalCount - 25

		# Add modified original row with COUNT = 1
		$newRow = $exportTable.NewRow()
		$newRow["CHARGE CODE"] = $originalCode
		$newRow["BILLCODE"]     = $agency
		$newRow["COUNT"]        = 1
		$exportTable.Rows.Add($newRow)

		# Collect 1232A data if adjusted > 0
		if ($adjusted -gt 0) {
			if ($dict1232A.ContainsKey($agency)) {
				$dict1232A[$agency] += $adjusted
			} else {
				$dict1232A[$agency] = $adjusted
			}
		}
	}

	# Step 2: Add aggregated 1232A rows
	foreach ($agency in $dict1232A.Keys) {
		$aRow = $exportTable.NewRow()
		$aRow["CHARGE CODE"] = "1232A"
		$aRow["BILLCODE"]     = $agency
		$aRow["COUNT"]        = $dict1232A[$agency]
		$exportTable.Rows.Add($aRow)
	}

	# Step 3: Output file name with dynamic month
	$monthName = (Get-Date).ToString("MMMM")
	$outputFile = "D:\Scripts\${monthName} - 1232,1232A,1232C Billing.csv"

	# Export to CSV
	$exportTable | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
	Write-Host "Data exported successfully to $outputFile"	
} catch {
    Write-Error ("Database Exception: {0}`n{1}" -f $con.ConnectionString, $_.Exception.ToString())
} finally {
    if ($con.State -eq 'Open') { $con.Close() }
}
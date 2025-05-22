$leftFile = Read-Host "Enter the path of the first csv, the diff will show what is in the first csv and not in the second"
$rightFile = Read-Host "Enter the full path of the second csv"
$OutputFile = Read-Host "Enter the path to save the diff, leave off a file name"

if ($OutputFile.Substring($OutputFile.Length - 1) -ne "\") {
    $OutputFile += "\"
}

$csv1 = Import-Csv "$leftFile"
$csv2 = Import-Csv "$rightFile"

$file1Values = $csv1.Name
$file2Values = $csv2.Name

$diff = $file1Values | Where-Object {$file2Values -notcontains $_}

$diff | ForEach-Object { [PSCustomObject]@{ Column1 = $_ } } | Export-Csv "$OutputFile\diff.csv" -NoTypeInformation
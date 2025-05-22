$path1 = Read-Host "Enter the path of where you are copying from"
$path2 = Read-Host "Enter the path of wher you are copying to"

if ($path1.Substring($path1.Length - 1) -ne "\") {
    $path1 += "\"
}

if ($path2.Substring($path2.Length - 1) -ne "\") {
    $path2 += "\"
}


$foldersInLeft = Get-ChildItem $path1 -Directory | Select-Object -Property Name
$foldersInRight = Get-ChildItem $path2 -Directory | Select-Object -Property Name

$file1Values = $foldersInLeft.Name
$file2Values = $foldersInRight.Name

$diff = $file1Values | Where-Object {$file2Values -notcontains $_}

$folders = $diff | ForEach-Object { [PSCustomObject]@{ Column1 = $_ } } | Export-Csv -NoTypeInformation -Path $null

foreach ($folder in $folders) {
    $src = "$path1$($folder.Column1)"
    $dst = "$path2\$($folder.Column1)"

    if (Test-Path $src) {
        Write-Host "Copying from source folder: $src to destination folder: $dst"
        Copy-Item $src $dst -Recurse
    } else {
        Write-Host "Source folder $src not found."
    }
}

Read-Host -Prompt "Press Enter to close the window"
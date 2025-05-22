function Convert-DecimalToBase36 {
  param (
    [int] $decimalNumber
  )

  $base36 = ''
  $alphabet = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'

  while ($decimalNumber -gt 0) {
    $remainder = $decimalNumber % 36
    $base36 = $alphabet[$remainder] + $base36
    $decimalNumber = [math]::Floor($decimalNumber / 36)
  }

  return $base36.PadLeft(8, '0')
}

$inputFile = Read-Host "Enter the path of the input file"

$fileName = [System.IO.Path]::GetFileNameWithoutExtension($inputFile)
$extension = [System.IO.Path]::GetExtension($inputFile)
$outputFile = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($inputFile), "$fileName-converted$extension")

$decimalNumbers = Get-Content $inputFile

$base36Numbers = @()
foreach ($decimalNumber in $decimalNumbers) {
  $base36Numbers += Convert-DecimalToBase36 $decimalNumber
}

$base36Numbers | Out-File $outputFile

Write-Output "Output file is located at $outputFile"
Read-Host -Prompt "Press Enter to close this window"
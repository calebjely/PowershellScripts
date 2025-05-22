# Convert-HDA-Files-to-CSV

What it does:
1. Finds all .hda files in the specified folder
2. Processes them in alphabetical order
3. Includes column headers from the first file only
4. Combines all records into a single CSV
5. Saves the output in the same folder
6. Reports total files processed, columns, and records

Parameters:
FolderPath: UNC path to the folder containing HDA files (required)
OutputFileName: Name of the output CSV file (optional, defaults to "combined_export.csv")

Usage:
.\Convert-HDAToCSV.ps1 -FolderPath "\\server\share\folder"
.\Convert-HDAToCSV.ps1 -FolderPath "C:\path\to\hda\files" -OutputFileName "my_export.csv"


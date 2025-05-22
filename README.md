# ConvertHDAToCSV
### What it does:
1. Finds all .hda files in the specified folder
2. Processes them in alphabetical order
3. Includes column headers from the first file only
4. Combines all records into a single CSV
5. Saves the output in the same folder
6. Reports total files processed, columns, and records

### Parameters:
FolderPath: UNC path to the folder containing HDA files (required)
OutputFileName: Name of the output CSV file (optional, defaults to "combined_export.csv")

### Usage:
.\Convert-HDAToCSV.ps1 -FolderPath "\\server\share\folder"
.\Convert-HDAToCSV.ps1 -FolderPath "C:\path\to\hda\files" -OutputFileName "my_export.csv"

# ConvertDecimalListToFolderName
Input a file with a list of ODC Batch Ids, Produce a list of Folder patchs to those batches.

# GetDiffBetweenTwoCsvs
Input 2 csv files and procduce a csv file of the diff between the two files.

# MoveFoldersFromSourceToDestination.ps1
Input a source and destination path. Moves files from the source to the destination.

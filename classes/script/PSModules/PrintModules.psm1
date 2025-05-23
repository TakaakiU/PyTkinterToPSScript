enum XlLookAt {
    xlWhole = 1
    xlPart = 2
}
enum XlFixedFormatType {
    xlTypePDF = 0
    xlTypeXPS = 1
}

# Load the class to be used
$packageListConfig = "$(($PSScriptRoot).Replace('\', '/'))/../PSClasses/PrintListConfig.ps1"
if (-not (Test-Path $packageListConfig)) {
    Write-Error "One or more psm files do not exist when loading class modules.: $($_.Exception.Message)"
}
. $packageListConfig
[System.String]$TEMPLATESHEET1 = [PrintListConfig]::TEMPLATESHEET1
[System.Int32]$TEMPLATEROWS1 = [PrintListConfig]::TEMPLATEROWS1
[System.String]$TEMPLATESHEET2 = [PrintListConfig]::TEMPLATESHEET2
[System.Int32]$TEMPLATEROWS2 = [PrintListConfig]::TEMPLATEROWS2
[System.String]$HEADERRANGE = [PrintListConfig]::HEADERRANGE
[System.String]$MAINRANGE = [PrintListConfig]::MAINRANGE

# Individual functions
Function Test-ExcelSheetExists {
    param(
        [System.String]$Path,
        [System.String]$CheckSheet
    )

    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    $sheetExists = $false
    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
        $workSheets = $workBook.Sheets

        foreach ($workSheet in $workSheets) {
            if ($workSheet.Name -eq $CheckSheet) {
                $sheetExists = $true
                break
            }
        }
    }
    catch {
        Write-Error 'An error occurred during Excel operation.'
    }
    finally {
        # Release up to workbook
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # Close without saving the workbook
            $workBook.Close($false)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Exit Excel application
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }

    return $sheetExists
}

# Check if the file is locked
function Test-FileLocked {
    param (
        [Parameter(Mandatory=$true)][System.String]$Path
    )

    if (-Not(Test-Path $Path)) {
        Write-Error 'The target file does not exist.' -ErrorAction Stop
    }

    # Convert to absolute path because Open method does not work properly with relative paths
    $fullPath = (Resolve-Path -Path $Path -ErrorAction SilentlyContinue).Path

    $fileLocked = $false
    try {
        # Open the file in read-only mode
        $fileStream = [System.IO.File]::Open($fullPath, 'Open', 'ReadWrite', 'None')
    }
    catch {
        # If the file cannot be opened, consider it locked
        $fileLocked = $true
    }
    finally {
        if ($null -ne $fileStream) {
            $fileStream.Close()
        }
    }

    return $fileLocked
}

# Check if the sheet name is valid
function Test-ExcelSheetname {
    param(
        [Parameter(Mandatory=$true)][System.String]$SheetName
    )
    
    # Check if the name is not blank
    if ([string]::IsNullOrWhiteSpace($SheetName.Trim())) {
        Write-Warning 'The name is blank (including null or empty string)'
        return $false
    }

    # Check if the length is within 31 characters
    if ($SheetName.Length -gt 31) {
        Write-Warning 'The length is not within 31 characters'
        return $false
    }

    # Check if it contains prohibited characters
    if ($SheetName -match "[:\\\/\?\*\[\]]") {
        Write-Warning 'Contains colon (:), backslash (\), slash (/), question mark (?), asterisk (*), or square brackets ([])'
        return $false
    }

    return $true
}

# Check the file extension of a string
function Test-FileExtension {
    param (
        [Parameter(Mandatory=$true)][System.String]$FullFilename,
        [Parameter(Mandatory=$true)][System.String[]]$Extensions
    )

    # Check if the string exists
    #   Check for null, empty, or whitespace
    if ([System.String]::IsNullOrWhiteSpace($FullFilename.Trim())) {
        Write-Error 'The string to be checked is not set.'
        return $false
    }
    #   Check if it contains a period
    if ($FullFilename -notmatch '\.') {
        Write-Error 'The string to be checked does not contain a period.'
        return $false
    }
    #   Check if the period is not at the beginning or end
    $dotIndex = $FullFilename.LastIndexOf('.')
    if (($dotIndex -eq 0) -or
        ($dotIndex -eq $FullFilename.Length - 1)) {
        Write-Error 'The string to be checked is not a valid file name.'
        return $false
    }

    # Check within the array
    foreach ($item in $Extensions) {
        # Check for null, empty, or whitespace
        if ([System.String]::IsNullOrWhiteSpace($item.Trim())) {
            Write-Warning 'There is data in the extension array that is not set.'
            return $false
        }
        # Check if the first character starts with a period
        if ($item -notmatch '^\.') {
            Write-Warning 'There is data in the extension array that does not start with a period.'
            return $false
        }
    }

    # Check the extension
    #   Get the extension
    [System.String]$fileExtension = $FullFilename -replace '.*(\..*)', '$1'

    #   Compare the extension
    $isHit = $false
    foreach ($item in $Extensions) {
        # If the extension matches
        if ($fileExtension -eq $item) {
            $isHit = $true
            break
        }
    }

    # Return the result
    return $isHit
}

# Check if the sheet exists
Function Test-ExcelSheetExists {
    param(
        [System.String]$Path,
        [System.String]$CheckSheet
    )

    $sheetExists = $false

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
        $workSheets = $workBook.Sheets

        foreach ($workSheet in $workSheets) {
            if ($workSheet.Name -eq $CheckSheet) {
                $sheetExists = $true
                break
            }
        }
    }
    catch {
        Write-Error "Unexpected error occurred during sheet existence check. [Details: $($_.Exception.Messsage)]"
    }
    finally {
        # Release up to workbook
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # Close without saving the workbook
            $workBook.Close($false)
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Exit Excel application
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }

    return $sheetExists
}

# Function to delete specified sheet names
function Remove-ExcelSheets {
    param (
        [System.String]$Path,
        [System.String[]]$RemoveSheets
    )

    # Input validation
    if (-not (Test-Path $Path)) {
        Write-Warning "The target path is not valid. [Target path: $($Path)]"
        return
    }
    # Check the extension
    elseif (-not (Test-FileExtension $Path @('.xls', '.xlsx'))) {
        Write-Warning "The file in the target path is not an Excel file. [Target path: $($Path)]"
        return
    }
    # Check the file lock state
    elseif (Test-FileLocked($Path)) {
        Write-Warning "The target file is open. Close the file and try again. [Target file: $($Path)]"
        return
    }

    # Sheet deletion process
    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    try {
        # Reference COM object
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false

        # Open the target file
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)
                
        # Reference sheets
        $workSheets = $workBook.Worksheets

        # Iterate through the sheets specified in the argument
        foreach ($removeSheet in $RemoveSheets) {
            # If there is only one sheet, interrupt the deletion process
            if ($workSheets.Count -eq 1){
                Write-Warning "Currently, there is only one sheet in Excel. Since Excel requires at least one sheet, the deletion process is interrupted."
                break
            }
            # Check the value of the sheet name (empty string, valid sheet name)
            elseif (-not (Test-ExcelSheetname $removeSheet)) {
                Write-Warning "The sheet name to be deleted is not valid. [Sheet name to be deleted: $($removeSheet)]]"
                continue
            }
            # Check the existence of the sheet to be deleted
            elseif (-not (Test-ExcelSheetExists $Path $removeSheet)) {
                Write-Warning "The sheet name to be deleted does not exist. [Target path: $($Path), Sheet name to be deleted: $($removeSheet)]]"
                continue
            }

            # Delete the sheet
            $workSheet = $workSheets.Item($removeSheet)
            $workSheet.Delete()
        }
    }

    catch {
        Write-Error "Unexpected error occurred during sheet deletion process. [Details: $($_.Exception.Message), Location: $($_.InvocationInfo.MyCommand.Name)]"
    }

    finally {
        # Release up to workbook
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # Save and exit
            $workBook.Close($true)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Exit Excel application
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}

# Determine the type of array
function Get-ArrayType {
    param(
        $InputObject
    )
    
    [System.Collections.Hashtable]$arrayTypes = @{
        "OtherTypes" = -1
        "SingleArray" = 0
        "MultiLevel" = 1
        "MultiDimensional" = 2
    }

    # If there is no data
    if ($null -eq $InputObject) {
        return $arrayTypes["OtherTypes"]
    }

    # If the outermost is not an array
    if ($InputObject -isnot [System.Array]) {
        return $arrayTypes["OtherTypes"]
    }

    # Determine if it is a jagged array (multi-level array)
    $isMultiLevel = $false
    foreach ($element in $InputObject) {
        if ($element -is [System.Array]) {
            # Multi-level array where the array inside is also an array
            $isMultiLevel = $true
            break
        }
    }
    if ($isMultiLevel) {
        return $arrayTypes["MultiLevel"]
    }    
    
    # Determine if it is a multi-dimensional array
    if ($InputObject.Rank -ge 2) {
        # If it is 2-dimensional or higher
        return $arrayTypes["MultiDimensional"]
    }
    else {
        # If it is 1-dimensional
        # Assumption: It has already been confirmed to be an array by the initial "-isnot [System.Array]" check.
        return $arrayTypes["SingleArray"]
    }
}

# Compare multi-dimensional arrays
function Test-ArrayEquality {
    param (
        [Parameter(Mandatory=$true)]$Array1,
        [Parameter(Mandatory=$true)]$Array2
    )

    # Determine if it is a multi-dimensional array
    $resultArrayType = (Get-ArrayType $Array1)
    if ($resultArrayType -ne 2) {
        Write-Warning "The argument 'Array1' is not a multi-dimensional array. [Array type determination result: $($resultArrayType)]"
        return
    }
    $resultArrayType = (Get-ArrayType $Array2)
    if ($resultArrayType -ne 2) {
        Write-Warning "The argument 'Array2' is not a multi-dimensional array. [Array type determination result: $($resultArrayType)]"
        return
    }

    # Compare the number of dimensions in the arrays
    $dimensionArray1 = $Array1.Rank
    $dimensionArray2 = $Array2.Rank

    if ($dimensionArray1 -ne $dimensionArray2) {
        return $false
    }

    # Check the number of elements in each dimension
    for ($i = 0; $i -lt $dimensionArray1; $i++) {
        if ($Array1.GetLength($i) -ne $Array2.GetLength($i)) {
            return $false
        }
    }

    # The number of elements matches
    return $true
}

# Calculate the maximum number of pages
function Get-MaxPage {
    param(
        [System.Int32]$FirstPageCount = $TEMPLATEROWS1,
        [System.Int32]$OtherPageCount = $TEMPLATEROWS2,
        [System.Int32]$DataCount
    )

    if ($DataCount -le $FirstPageCount) {
        $maxPage = 1
    } else {
        $maxPage = [System.Math]::Ceiling(($dataCount - $firstPageCount) / $otherPageCount)
        $maxPage += 1
    }
    return $maxPage
}

# Set values in the sheet using constants as keys
function Set-PrintListValues {
    param(
        [PrintListConfig]$Config
    )

    # Get item names
    $headerColumns = @($Config.HeaderConstants[0].psobject.properties | ForEach-Object { $_.Name })
    $mainColumns = @($Config.MainConstants[0].psobject.properties | ForEach-Object { $_.Name })


    # Replacement process
    $excelApp = $null
    $workBooks = $null
    $workBook = $null
    $workSheets = $null
    $workSheet = $null

    # Current page
    $currentPage = 1
    $maxPage = (Get-MaxPage -DataCount $Config.MainValues.Count)
    $currentRow = 1

    try {
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Config.Path)
        $workSheets = $workBook.Sheets

        # First page
        #   Prepare sheet
        $currentPage = 1
        $currentSheet = "$($currentPage) Page"
        $workSheet = $workSheets.Item($TEMPLATESHEET1)
        # Copy sheet
        $workSheet.Copy($workSheet)
        # Change the name of the copied sheet (assuming the copied sheet becomes the active sheet)
        $excelApp.ActiveSheet.Name = $currentSheet

        #   Reflect header information
        $workSheet = $workSheets.Item($currentSheet)
        $range = $workSheet.Range($HEADERRANGE)
                
        for ($i = 0; $i -lt $Config.HeaderConstants.Count; $i++) {
            for ($j = 0; $j -lt $headerColumns.Count; $j++) {
                $range.Replace(
                    $($Config.HeaderConstants[$i]."$($headerColumns[$j])"),
                    $($Config.HeaderValues[$i]."$($headerColumns[$j])"),
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        #   Reflect main information
        $startRow = $currentRow - 1
        [System.Int32]$rangeRows = $TEMPLATEROWS1
        $rangeMainValues = @($Config.MainValues | Select-Object -Skip $startRow -First $rangeRows)

        $workSheet = $workSheets.Item($currentSheet)
        $range = $workSheet.Range($MAINRANGE)

        for ($i = 0; $i -lt $rangeMainValues.Count; $i++) {
            # For constants with serial numbers
            $num = "{0:D2}" -f $($i + 1)
            for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                # Add serial number to constant
                $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                $range.Replace(
                    $mainConstantsWithNum,
                    $($rangeMainValues[$i]."$($mainColumns[$j])"),
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        #   Create blanks
        for ($i = 0; $i -lt $TEMPLATEROWS1; $i++) {
            # For constants with serial numbers
            $num = "{0:D2}" -f $($i + 1)
            for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                # Add serial number to constant
                $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                $range.Replace(
                    $mainConstantsWithNum,
                    '',
                    # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                    [XlLookAt]::xlWhole
                ) | Out-Null
            }
        }

        # Move to the next page
        $currentRow = $TEMPLATEROWS1 + 1

        # From the second page onwards
        for ($currentPage = 2; $currentPage -le $maxPage; $currentPage++) {
            #   Prepare sheet
            $currentSheet = "$($currentPage) Page"
            $workSheet = $workSheets.Item($TEMPLATESHEET2)
            # Copy sheet
            $workSheet.Copy($workSheet)
            # Change the name of the copied sheet (assuming the copied sheet becomes the active sheet)
            $excelApp.ActiveSheet.Name = $currentSheet

            #   Reflect header information
            $workSheet = $workSheets.Item($currentSheet)
            $range = $workSheet.Range($HEADERRANGE)

            for ($i = 0; $i -lt $Config.HeaderConstants.Count; $i++) {
                for ($j = 0; $j -lt $headerColumns.Count; $j++) {
                    $range.Replace(
                        $($Config.HeaderConstants[$i]."$($headerColumns[$j])"),
                        $($Config.HeaderValues[$i]."$($headerColumns[$j])"),
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            #   Reflect main information
            $startRow = $currentRow - 1
            [System.Int32]$rangeRows = $TEMPLATEROWS2
            $rangeMainValues = @($Config.MainValues | Select-Object -Skip $startRow -First $rangeRows)

            $workSheet = $workSheets.Item($currentSheet)
            $range = $workSheet.Range($MAINRANGE)

            for ($i = 0; $i -lt $rangeMainValues.Count; $i++) {
                # For constants with serial numbers
                $num = "{0:D2}" -f $($i + 1)
                for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                    # Add serial number to constant
                    $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                    $range.Replace(
                        $mainConstantsWithNum,
                        $($rangeMainValues[$i]."$($mainColumns[$j])"),
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            #   Create blanks
            for ($i = 0; $i -lt $TEMPLATEROWS2; $i++) {
                # For constants with serial numbers
                $num = "{0:D2}" -f $($i + 1)
                for ($j = 0; $j -lt $mainColumns.Count; $j++) {
                    # Add serial number to constant
                    $mainConstantsWithNum = $($Config.MainConstants[0]."$($mainColumns[$j])") + $num
                    $range.Replace(
                        $mainConstantsWithNum,
                        '',
                        # [Microsoft.Office.Interop.Excel.XlLookAt]::xlWhole
                        [XlLookAt]::xlWhole
                    ) | Out-Null
                }
            }

            # Move to the next page
            $currentRow += $TEMPLATEROWS2
        }
    }
    catch {
        # Error handling
        Write-Error "Unexpected error occurred. [Details: $($_.Exception.ToString())]"
    }
    finally {
        # Release up to workbook
        if ($null -ne $workSheet) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheet) | Out-Null
            $workSheet = $null
            Remove-Variable workSheet -ErrorAction SilentlyContinue
        }
        if ($null -ne $workSheets) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workSheets) | Out-Null
            $workSheets = $null
            Remove-Variable workSheets -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBook) {
            # Save and exit workbook
            $workBook.Close($true)
                    
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Exit Excel application
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}

function Export-ExcelDocumentAsPDF {
    param(
        [parameter(Mandatory=$true)][string]$Path,
        [parameter(Mandatory=$true)][string]$OutputPath
    )

    $excelApp = $null
    $workBooks = $null
    $workBook = $null

    try {
        # Open Excel file
        $excelApp = New-Object -ComObject Excel.Application
        $excelApp.Visible = $false
        $excelApp.DisplayAlerts = $false
        $workBooks = $excelApp.Workbooks
        $workBook = $workBooks.Open($Path)

        # Export as PDF
        $workBook.ExportAsFixedFormat(
            # [Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF,
            [XlFixedFormatType]::xlTypePDF,
            $OutputPath
        )
    }
    catch {
        Write-Error "Error occurred while exporting Excel file as PDF."
        Write-Error "Error details [Details: $($_.Exception.Message), Location: $($_.InvocationInfo.MyCommand.Name)]"
    }
    finally {
        if ($null -ne $workBook) {
            # Close workbook without saving
            $workBook.Close($false)
                    
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBook) | Out-Null
            $workBook = $null
            Remove-Variable workBook -ErrorAction SilentlyContinue
        }
        if ($null -ne $workBooks) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workBooks) | Out-Null
            $workBooks = $null
            Remove-Variable workBooks -ErrorAction SilentlyContinue
        }

        # Exit Excel application
        if ($null -ne $excelApp) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()

            $excelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
            $excelApp = $null
            Remove-Variable excelApp -ErrorAction SilentlyContinue

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
    }
}

# Determine if it is a PSCustomObject
function Test-IsPSCustomObject {
    param(
        [Parameter(Mandatory=$true)]
        [System.Object[]]$Argument
    )

    foreach ($arg in $Argument) {
        if (-not ($arg -is [System.Management.Automation.PSCustomObject])) {
            return $false
        }
    }
    return $true
}

# Compare PSCustomObjects
Function Test-PSCustomObjectEquality {
    param (
        [Parameter(Mandatory=$true)][System.Object[]]$Object1,
        [Parameter(Mandatory=$true)][System.Object[]]$Object2
    )

    # Check if data exists
    if (($Object1.Count -eq 0) -or ($Object2.Count -eq 0)) {
        return $false
    }

    # Determine if the objects are PSCustomObjects
    if (-not (Test-IsPSCustomObject $Object1)) {
        return $false
    }
    elseif (-not (Test-IsPSCustomObject $Object2)) {
        return $false
    }

    # Compare item names
    $object1ColumnData = $Object1[0].psobject.properties | ForEach-Object { $_.Name }
    $object2ColumnData = $Object2[0].psobject.properties | ForEach-Object { $_.Name }
    $compareResult = (Compare-Object $object1ColumnData $object2ColumnData -SyncWindow 0)
    if (($null -ne $compareResult) -and ($compareResult.Count -ne 0)) {
        return $false
    }

    # The two objects match
    return $true
}

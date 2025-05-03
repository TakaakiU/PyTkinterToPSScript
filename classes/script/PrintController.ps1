param (
    [Switch]$FormA,
    [Switch]$FormB,
    [System.String]$OutputPath,
    [System.String]$RootPath = 'C:/PyTkinterToPSScript',
    [System.String]$DataMapping_Header = 'DataMapping_Header.csv',
    [System.String]$DataMapping_Body = 'DataMapping_Body.csv',
    [System.String]$FormA_Template = 'Template_FormA.xlsx',
    [System.String]$FormA_HeaderValues = 'FormA_HeaderValues.csv',
    [System.String]$FormA_BodyValues = 'FormA_BodyValues.csv',
    [System.String]$FormB_Template = 'Template_FormB.xlsx',
    [System.String]$FormB_HeaderValues = 'FormB_HeaderValues.csv',
    [System.String]$FormB_BodyValues = 'FormB_BodyValues.csv'
)

# Main process
$statusCode = 0

# Load the required class
$packageListConfig = "$(($PSScriptRoot).Replace('\', '/'))/PSClasses/PrintListConfig.ps1"
# Load the required modules
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$printModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/PrintModules.psm1"
if (-not (Test-Path $packageListConfig) -or
    -not (Test-Path $commonModules) -or
    -not (Test-Path $printModules)) {
    $statusCode = -8001
    Write-Error "Required external module files (*.ps1, *.psm1) are missing.: $($_.Exception.Message)"
}
else {
    try {
        # Load custom class
        # Import-Module $packageListConfig -Force
        . $packageListConfig
        [System.String]$TEMPLATESHEET1 = [PrintListConfig]::TEMPLATESHEET1
        [System.String]$TEMPLATESHEET2 = [PrintListConfig]::TEMPLATESHEET2

        # Load common functions
        Import-Module $commonModules -Force
        
        # Load functions used for Pack
        Import-Module $printModules -Force
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        $statusCode = -8002
    }
}

# Logical check of arguments
if ($statusCode -eq 0) {
    if ($FormA -and $FormB) {
        Write-Error "Please review the arguments. [Reason: Both FormA and FormB are set]"
        $statusCode = -8003
    }
    elseif (-not $FormA -and -not $FormB) {
        Write-Error "Please review the arguments. [Reason: Neither FormA nor FormB is set]"
        $statusCode = -8004
    }
}

# Declare paths and check arguments
if ($statusCode -eq 0) {
    # Declare paths
    #   Template files
    $TemplatePath = "$($RootPath)/template"
    $formaPath = "$($TemplatePath)/$($FormA_Template)"
    $formbPath = "$($TemplatePath)/$($FormB_Template)"
    $headerMappingPath = "$($TemplatePath)/$($DataMapping_Header)"
    $bodyMappingPath = "$($TemplatePath)/$($DataMapping_Body)"
    #   Input files
    $InputPath = "$($RootPath)/input"
    $forma_HeaderValuesPath = "$($InputPath)/$($FormA_HeaderValues)"
    $forma_BodyValuesPath = "$($InputPath)/$($FormA_BodyValues)"
    $formb_HeaderValuesPath = "$($InputPath)/$($FormB_HeaderValues)"
    $formb_BodyValuesPath = "$($InputPath)/$($FormB_BodyValues)"

    # Check arguments
    if ($FormA) {
        $formPath = $formaPath
        $headerValuesPath = $forma_HeaderValuesPath
        $bodyValuesPath = $forma_BodyValuesPath
    }
    elseif ($FormB) {
        $formPath = $formbPath
        $headerValuesPath = $formb_HeaderValuesPath
        $bodyValuesPath = $formb_BodyValuesPath
    }
}

# Check the existence of folders and files
if ($statusCode -eq 0) {
    if (-not (Test-PathType $RootPath -PathRole Source -PathType Container)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $RootPath]"
        $statusCode = -8005
    }
    elseif (-not (Test-PathType $TemplatePath -PathRole Source -PathType Container)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $TemplatePath]"
        $statusCode = -8006
    }
    elseif (-not (Test-Path $formPath)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $formPath]"
        $statusCode = -8007
    }
    elseif (-not (Test-Path $headerMappingPath)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $headerMappingPath]"
        $statusCode = -8008
    }
    elseif (-not (Test-Path $bodyMappingPath)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $bodyMappingPath]"
        $statusCode = -8009
    }
    elseif (-not (Test-Path $headerValuesPath)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $headerValuesPath]"
        $statusCode = -8010
    }
    elseif (-not (Test-Path $bodyValuesPath)) {
        Write-Error "The folder or file specified in the arguments does not exist. [Target: $bodyValuesPath]"
        $statusCode = -8011
    }
}

# Check the existence of sheets in the Excel template file
if ($statusCode -eq 0) {
    if (-not (Test-ExcelSheetExists $formPath $TEMPLATESHEET1) -and
            -not (Test-ExcelSheetExists $formPath $TEMPLATESHEET2)) {
            Write-Error "The predefined sheets were not found in the Excel template file. [Target file: $($formPath), Sheet1: $([PrintListConfig]::TEMPLATESHEET1), Sheet2: $($TEMPLATESHEET2)]"
            $statusCode = -8012
    }
}

# Copy the Excel template file to a temporary folder
if ($statusCode -eq 0) {
    # C:/Users/XXX/AppData/Local/Temp/PyTkinterToPSScript_ExcelTemplaete.xlsx
    $excelWorkfilePath_format = "$(($Env:TEMP).Replace('\', '/'))/PyTkinterToPSScript_ExcelTemplaete.xlsx"

    # If the file already exists in the temporary folder, delete it
    try {
        Remove-File $excelWorkfilePath_format
        # If deletion succeeds, continue with the same file name
        $excelWorkfilePath = $excelWorkfilePath_format
    }
    catch {
        # If deletion fails for some reason, change the file name to a unique one
        $excelWorkfilePath = Get-UniqueFilePath $excelWorkfilePath_format
    }

    $copyFrom = $formPath
    $copyTo = $excelWorkfilePath
    try {
        Copy-Item $copyFrom $copyTo -Force
    }
    catch {
        Write-Error "An error occurred while copying the Excel template file. [Source: $($copyFrom), Destination: $($copyTo)]"
        $statusCode = -8013
    }
}

# Read and check header information
if ($statusCode -eq 0) {
    # Read and check header information
    try {
        # Input position data for header information
        $headerMapping = @(Import-Csv -Path $headerMappingPath)
        # Input data for header information
        $headerValue = @(Import-Csv -Path $headerValuesPath)
    }
    catch {
        Write-Error "An error occurred while reading the position data and input data for header information."
        $statusCode = -8014
    }

    if ($statusCode -eq 0) {
        if (-not (Test-PSCustomObjectEquality $headerMapping $headerValue)) {
            Write-Error 'The field names in the position data and input data for header information do not match.'
            $statusCode = -8015
        }
    }
}

# Read and check main data
if ($statusCode -eq 0) {
    try {
        # Input position data for header information
        $headerMapping = @(Import-Csv -Path $headerMappingPath)
        # Input data for header information
        $headerValue = @(Import-Csv -Path $headerValuesPath)
    }
    catch {
        Write-Error "An error occurred while reading the position data and input data for header information."
        $statusCode = -8016
    }

    if ($statusCode -eq 0) {
        if (-not (Test-PSCustomObjectEquality $headerMapping $headerValue)) {
            Write-Error 'The field names in the position data and input data for header information do not match.'
            $statusCode = -8017
        }
    }
}

# Read and check main data
if ($statusCode -eq 0) {
    try {
        # Input position data for main data
        $bodyMapping = @(Import-Csv -Path $bodyMappingPath)
        # Input data for main data
        $bodyValue = @(Import-Csv -Path $bodyValuesPath)
    }
    catch {
        Write-Error "An error occurred while reading the position data and input data for main data."
        $statusCode = -8018
    }

    if ($statusCode -eq 0) {
        if (-not (Test-PSCustomObjectEquality $bodyMapping $bodyValue)) {
            Write-Error 'The field names in the position data and input data for main data do not match.'
            $statusCode = -8019
        }
    }
}

# Set values in the Excel file
if ($statusCode -eq 0) {
    $config = [PrintListConfig]::new($excelWorkfilePath, $headerMapping, $headerValue, $bodyMapping, $bodyValue)
    try {
        Set-PrintListValues -Config $config
    }
    catch {
        Write-Error "An error occurred while setting input data in the Excel file. [Target file: $($excelWorkfilePath)]"
        Write-Error "Error details [Message: $($_.Exception.Message), Location: $($_.InvocationInfo.MyCommand.Name)]"
        $statusCode = -8020
    }
}

# Printing process
if ($statusCode -eq 0) {
    # Delete template sheets from the Excel file
    try {
        $removeSheet = @($TEMPLATESHEET1, $TEMPLATESHEET2)
        Remove-ExcelSheets $excelWorkfilePath $removeSheet
    }
    catch {
        Write-Error "An error occurred while deleting template sheets from the Excel file. [Target file: $($excelWorkfilePath), Target sheets: $($removeSheet -join ', ')]"
        $statusCode = -8021
    }
}

# Output PDF file
if ($statusCode -eq 0) {
    try {
        Remove-Item $OutputPath -Force
        Export-ExcelDocumentAsPDF $excelWorkfilePath $OutputPath
    }
    catch {
        Write-Error "An error occurred while outputting the PDF file. [Target file: $($excelWorkfilePath), Output destination: $($OutputPath)]"
        $statusCode = -8022
    }
}

exit $statusCode

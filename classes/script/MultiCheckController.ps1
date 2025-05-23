param (
    [System.String]$InputPath,
    [System.String]$OutputPath
)

$statusCode = 0

# Load the functions to be used
$adpackController = "$(($PSScriptRoot).Replace('\', '/'))/AdpackController.ps1"
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$adpackModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/adpackModules.psm1"
if (-not (Test-Path $adpackController) -or -not (Test-Path $commonModules) -or -not (Test-Path $adpackModules)) {
    Write-Error "Required external module files (*.ps1, *.psm1) do not exist."
    $statusCode = -7201
}
else {
    try {
        # Load functions for common use
        Import-Module $commonModules
        # Load functions for Pack use
        Import-Module $adpackModules
    }
    catch {
        Write-Error "An error occurred while loading external module files (*.ps1, *.psm1): $($_.Exception.Message)"
        $statusCode = -7202
    }
}

# Logical check for each argument
if ($statusCode -eq 0) {
    # Check if the input data path is a folder
    if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Container)) {
        Write-Error "Please specify the input data as a 'folder'."
        $statusCode = -7203
    }
}

# Check PowerShell version
if ($statusCode -eq 0) {
    $statusCode = Test-PowerShellVersion7OrLater
}

# Check for the existence of the adpack executable file
if ($statusCode -eq 0) {
    $exePath = "$(($PSScriptRoot).Replace('\', '/'))/../exe/adpack.exe"
    if (-not (Test-Path $exePath)) {
        Write-Error "The adpack executable file does not exist."
        $statusCode = -7204
    }
}

# Execute based on arguments
if ($statusCode -eq 0) {
    # Check for the existence of ZIP files and create a ZIP file list
    $zipFiles = Get-ChildItem -Path $InputPath -Filter "*.zip" | Select-Object -ExpandProperty FullName
    if (-not $zipFiles) {
        Write-Error "No ZIP files exist in the specified folder."
        $statusCode = -7205
    }
}

# Repeat checks based on the ZIP file list
if ($statusCode -eq 0) {
    # Repeatedly perform UnPack using AdpackController.ps1
    $results = @()
    foreach ($zipFile in $zipFiles) {
        Write-Host "Checking file: $zipFile"
        $process = Start-Process -FilePath "pwsh" -ArgumentList "-File $adpackController -Unpack -InputPath `"$zipFile`"" -NoNewWindow -Wait -PassThru
        # $process = Start-Process -FilePath "powershell" -ArgumentList "-File $adpackController -Unpack -InputPath `"$zipFile`"" -NoNewWindow -Wait -PassThru
        $statusCode = $($process.ExitCode)

        $results += [PSCustomObject]@{
            FilePath = $zipFile
            StatusCode = $statusCode
        }

        # Abort if an error occurs
        if ($statusCode -ne 0) {
            break
        }
    }

    # Output the list externally
    $results | Export-Csv -Path $OutputPath -Encoding UTF8 -NoTypeInformation
}

if ($statusCode -eq 0) {
    Write-Host "Completed successfully."
}
else {
    Write-Host "Completed with errors. [Result Code:$statusCode]"
}

exit $statusCode

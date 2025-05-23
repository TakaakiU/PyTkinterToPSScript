param (
    [Switch]$Pack,
    [Switch]$UnPack,
    [System.String]$Hash = "s256",  # Use adpack
    # [System.String]$Hash = "SHA256",  # Use custom function if necessary
    [System.String]$InputPath,
    [System.String]$OutputPath = "",
    [Switch]$Check,
    [Switch]$NoCheck
)

$statusCode = 0

# Load required functions
$commonModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/CommonModules.psm1"
$adpackModules = "$(($PSScriptRoot).Replace('\', '/'))/PSModules/AdpackModules.psm1"
if (-not (Test-Path $commonModules) -or -not (Test-Path $adpackModules)) {
    Write-Error "Required external module files (*.psm1) are missing."
    $statusCode = -7001
}
else {
    try {
        # Load common functions
        Import-Module $commonModules
        # Load functions for Pack
        Import-Module $adpackModules
    }
    catch {
        Write-Error "An error occurred while loading external module files (*.psm1): $($_.Exception.Message)"
        $statusCode = -7002
    }
}

# Logical check of arguments
if ($statusCode -eq 0) {
    if ($Pack -and $UnPack) {
        Write-Error "Please review the arguments. [Reason: Both Pack and UnPack are set]"
        $statusCode = -7003
    }
    elseif (-not $Pack -and -not $UnPack) {
        Write-Error "Please review the arguments. [Reason: Neither Pack nor UnPack is set]"
        $statusCode = -7004
    }
    elseif ($Pack -and ($Check -or $NoCheck)) {
        Write-Error "Please review the arguments. [Reason: Pack + Check or Pack + UnCheck is set]"
        $statusCode = -7005
    }
    elseif ($UnPack -and ($Check -and $NoCheck)) {
        Write-Error "Please review the arguments. [Reason: UnPack + Check + UnCheck are all set]"
        $statusCode = -7006
    }
}

# Logical check for each argument
if ($statusCode -eq 0) {
    if ($Pack) {
        # Check if the input path is a folder
        if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Container)) {
            Write-Error "For packing, specify a folder as the input data."
            $statusCode = -7007
        }
        else {
            # If the output path is not specified, automatically set it based on the input data
            if ($OutputPath -eq "") {
                $OutputPath = "$($InputPath).zip"
            }
            # Check if the output path is a ZIP file
            else {
                if (-not (Test-PathType -Path $OutputPath -PathRole Target -PathType Leaf -Extension zip)) {
                    Write-Error "For packing, specify a file (*.zip) as the output data."
                    $statusCode = -7008
                }
            }
        }
    }
    elseif ($UnPack) {
        # Check if the input path is a ZIP file
        if (-not (Test-PathType -Path $InputPath -PathRole Source -PathType Leaf -Extension zip)) {
            Write-Error "For unpacking, specify a file (*.zip) as the input data."
            $statusCode = -7009
        }
        if ($statusCode -eq 0) {
            # If the output path is not specified, automatically set it based on the input data
            if ($OutputPath -eq "") {
                # Set the folder containing the input data as the output folder
                $OutputPath = (Split-Path -Path $InputPath)
            }
            # Check if the output path is a folder
            elseif (-not (Test-PathType -Path $OutputPath -PathRole Target -PathType Container)) {
                Write-Error "For unpacking, specify a folder as the output data."
                $statusCode = -7010
            }
        }
    }
}

# Check PowerShell version
if ($statusCode -eq 0) {
    $statusCode = Test-PowerShellVersion7OrLater
}

# Check existence of adpack executable file
if ($statusCode -eq 0) {
    $exePath = "$(($PSScriptRoot).Replace('\', '/'))/../exe/adpack.exe"
    if (-not (Test-Path $exePath)) {
        Write-Error "The adpack executable file is missing."
        $statusCode = -7011
    }
}

# Execute based on arguments
if ($statusCode -eq 0) {
    # Execute for valid arguments
    if ($Pack) {
        $statusCode = (Compress-Package_Pack -ExePath $exePath -HashAlgorithm $Hash -FolderPath $InPutPath -ZipFilePath $OutputPath)
    }
    elseif ($UnPack) {
        # Unpack (decompress and check)
        if (!$Check -And !$NoCheck) {
            $statusCode = (Expand-Package_Adpack -ExePath $exePath -HashAlgorithm $Hash -ZipFilePath $InPutPath -FolderPath $OutputPath)
        }
        # Unpack + Check (decompress → check → delete, only check results)
        elseif ($Check) {
            $statusCode = (Expand-Package_Adpack -Check -ExePath $exePath  -HashAlgorithm $Hash -ZipFilePath $InPutPath -FolderPath $OutputPath)
        }
        # Unpack + NoCheck (decompress)
        elseif ($NoCheck) {
            $statusCode = (Expand-Package_Adpack -NoCheck -ExePath $exePath -HashAlgorithm $Hash -ZipFilePath $InPutPath -FolderPath $OutputPath)
        }
    }
}

if ($statusCode -eq 0) {
    Write-Host "Completed successfully."
}
else {
    Write-Host "Completed with errors. [Result code:$statusCode]"
}

exit $statusCode

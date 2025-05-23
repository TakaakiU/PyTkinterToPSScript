# Determine if PowerShell 7 or later is being used (required for compressing large files into ZIP)
#   PowerShell 5 runs on .NET Framework, while PowerShell 7 runs on .NET Core.
#   PowerShell 7 is required due to performance and memory efficiency improvements in .NET Core.
function Test-PowerShellVersion7OrLater {
    $statusCode = 0
    # Retrieve the PowerShell version
    $currentVersion = $PSVersionTable.PSVersion
    $requiredMajorVersion = 7

    # Check the version
    if ($currentVersion.Major -lt $requiredMajorVersion) {
        Write-Host "This script must be run on PowerShell 7 or later." -ForegroundColor Red
        Write-Host "Current version: $currentVersion" -ForegroundColor Yellow
        # Abort processing
        $statusCode = -6001
    }

    return $statusCode
}

# Check file or folder
Function Test-PathType {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Source", "Target")]
        [string]$PathRole, # Role of the path: Input (Source) or Output (Target)

        [Parameter(Mandatory = $true)]
        [ValidateSet("Container", "Leaf")]
        [string]$PathType, # Expected path type: Folder (Container) or File (Leaf)
        
        [Parameter(Mandatory = $false)]
        [string[]]$Extensions = $null # Default: No extension check
    )

    $isMatch = $false
    
    # Argument check
    if (($PathType -eq "Container") -And $Extensions) {
        Write-Error "Invalid argument combination. Please review the arguments."
        return $isMatch
    }

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $localPathtype = $PathType
    # If "Target", get the parent directory
    if ($PathRole -eq "Target") {
        $parentDir = Split-Path -Path $fullPath
        if (-Not (Test-Path -Path $parentDir)) {
            Write-Error "The parent directory of the specified target path does not exist: $parentDir"
            return $isMatch
        }
        # Change the check target to the parent directory
        $fullPath = $parentDir
        $localPathtype = "Container"
    }

    # Check if the path exists
    if (-Not (Test-Path -Path $fullPath)) {
        Write-Error "The specified path does not exist: $fullPath"
        return $isMatch
    }

    # Check the path type
    switch ($localPathtype) {
        "Leaf" {
            if (-Not (Test-Path -Path $fullPath -PathType Leaf)) {
                Write-Error "A file was expected, but a folder was specified: $fullPath"
                return $isMatch
            }
        }
        "Container" {
            if (-Not (Test-Path -Path $fullPath -PathType Container)) {
                Write-Error "A folder was expected, but a file was specified: $fullPath"
                return $isMatch
            }
        }
    }

    # Check file extensions (for files)
    if ($Extensions -and $PathType -eq "Leaf") {
        $extension = [System.IO.Path]::GetExtension($Path).TrimStart(".").ToLower()
        if (-Not ($Extensions -contains $extension)) {
            Write-Error "The specified file extension '$extension' is not allowed: $Extensions"
            return $isMatch
        }
    }

    # Pass all validations
    $isMatch = $true
    return $isMatch
}

# Get a unique file path
function Get-UniqueFilePath {
    param (
        [System.String]$Path,
        [System.Int32]$MaxAttempts = 0
    )

    # Separate the file name and extension (no extension for folders)
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $Extension = [System.IO.Path]::GetExtension($Path)
    $ParentFolder = [System.IO.Path]::GetDirectoryName($Path)

    # If it doesn't exist, return as is.
    if (-not (Test-Path $Path)) {
            return $Path
    }

    # Set a unique path
    if ($MaxAttempts -eq 0) {
        $counter = 1
        do {
            # Generate a new path name (add sequential numbers)
            $newPath = "$ParentFolder\$BaseName-$counter$Extension"
            $counter++
        } while (Test-Path $newPath)  # Repeat until a non-existent path is found
    }
    else {
        for ($counter = 1; $counter -le $MaxAttempts; $counter++) {
            $newPath = "$ParentFolder\$BaseName-$counter$Extension"

            if (-not (Test-Path $newPath)) {
                break
            }
        }

        if ($counter -gt $MaxAttempts) {
            throw "A unique path was obtained, but the number of attempts '$MaxAttempts' was exceeded, so the process is aborted."
        }
    }

    return $newPath
}

# Delete a folder
function Remove-Folder {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    if (Test-Path -Path $FolderPath) {
        try {
            Remove-Item -Path $FolderPath -Recurse -Force | Out-Null
        }
        catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
            $statusCode = -6101
        }
    }
    else {
        Write-Host "The folder to be deleted does not exist, so the process is skipped: [$FolderPath]"
    }

    return $statusCode
}

# Delete a file
function Remove-File {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FilePath
    )

    [System.Int32]$statusCode = 0

    if (Test-Path -Path $FilePath) {
        try {
            Remove-Item -Path $FilePath -Force | Out-Null
        }
        catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
            $statusCode = -6102
        }
    }
    else {
        Write-Host "The file to be deleted does not exist, so the process is skipped: [$FilePath]"
    }

    return $statusCode
}

# Function to compress a folder into a ZIP file
function Compress-FolderToZip {
    param (
        [System.String]$FolderPath,
        [System.String]$Destination
    )
    [System.Int32]$statusCode = 0

    # If a ZIP file already exists, delete it
    if (Test-Path $Destination) {
        Remove-Item -Path $Destination -Force
    }    
    # Compression
    try {
        # ([System.IO.Compression.ZipFile]::CreateFromDirectory($FolderPath, $Destination)) | Out-Null
        # Compress-Archive -Path $FolderPath -DestinationPath $Destination | Out-Null

        # 7Zip4Powershell
        Compress-7Zip -Path $FolderPath -OutputPath $Destination
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        $statusCode = -6201
    }

    return $statusCode
}

# Function to extract a ZIP file into a temporary folder
function Expand-ZipToTempFolder {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$ZipFilePath,
        [Parameter(Mandatory = $true)]
        [System.String]$TempFolderPath
    )
    [System.Int32]$statusCode = 0

    try {
        # Extract the ZIP file
        # Expand-Archive -Path $ZipFilePath -DestinationPath $TempExtractFolder | Out-Null
        $encodingSjis = [Text.Encoding]::GetEncoding("shift_jis")
        ([System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFilePath, $TempExtractFolder, $encodingSjis)) | Out-Null

        # # 7Zip4Powershell
        # Expand-7Zip -ArchiveFileName $ZipFilePath -TargetPath $TempExtractFolder
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        $statusCode = -6301
    }

    return $statusCode
}

# Create a folder (Specify the switch parameter "-ForceRecreate" to delete and create)
function New-DirectoryIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath,
        [Switch]$ForceRecreate
    )
    [System.Int32]$statusCode = 0

    if (-Not (Test-Path -Path $FolderPath)) {
        try {
            New-Item -ItemType Directory -Path $FolderPath | Out-Null
            Write-Debug "Folder created: $FolderPath"
        }
        catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
            $statusCode = -6401
        }
    } else {
        if ($ForceRecreate) {
            try {
                Remove-Folder -FolderPath $FolderPath | Out-Null
                New-Item -ItemType Directory -Path $FolderPath | Out-Null
                Write-Debug "Folder deleted and recreated: $FolderPath"
            }
            catch {
                Write-Error "An error occurred: $($_.Exception.Message)"
                $statusCode = -6402
            }
        } else {
            Write-Debug "Folder already exists: $FolderPath"
        }
    }

    return $statusCode
}

# Compare the contents of a ZIP file and a target folder
function Compare-ZipAndFolderContent {
    param (
        [Parameter(Mandatory = $true)]
        [System.String]$FolderPath,
    
        [Parameter(Mandatory = $true)]
        [System.String]$ZipFilePath
    )

    [System.Int32]$statusCode = 0

    # Create a folder for temporary storage for extraction
    $zipDirectory = (Split-Path -Path $ZipFilePath) -replace('\\', '/')
    $TempExtractFolder = $zipDirectory + "/.TempExtract_" + [System.Guid]::NewGuid()
    # If it exists, delete it before creating
    $statusCode = (New-DirectoryIfNotExists $TempExtractFolder -ForceRecreate)

    # Create a temporary folder for extraction
    $statusCode = (Expand-ZipToTempFolder $ZipFilePath $TempExtractFolder)
    
    # Compare
    if ($statusCode -eq 0) {
        # Get the directory structure of the input data (including system files)
        $SourceItems = Get-ChildItem -Path $FolderPath -Recurse -Force | ForEach-Object {$_.FullName -replace "\\", "/"}
        $SourceItems_DelString = $FolderPath + "/"
        $SourceItems = $SourceItems -replace $SourceItems_DelString, ""

        # Get the directory structure of the output data
        $ExtractedItems = Get-ChildItem -Path $TempExtractFolder -Recurse -Force | ForEach-Object {$_.FullName -replace "\\", "/"}
        $ExtractedItems_DelString = $TempExtractFolder + "/"
        $ExtractedItems = $ExtractedItems -replace $ExtractedItems_DelString, ""

        Write-Debug ($SourceItems | Out-String)
        Write-Debug ($ExtractedItems | Out-String)

        # Change the base and get the differences
        $OnlyInSource = $SourceItems | Where-Object { $_ -notin $ExtractedItems }
        $OnlyInZip = $ExtractedItems | Where-Object { $_ -notin $SourceItems }

        # Display comparison results
        if ($OnlyInSource.Count -eq 0 -and $OnlyInZip.Count -eq 0) {
            Write-Host "The contents of the folder and the ZIP file match."
        }
        else {
            if ($OnlyInSource.Count -gt 0) {
                Write-Host "The following items exist only in the folder [$FolderPath]:"
                Write-Host ($OnlyInSource | Out-String)
            }
            if ($OnlyInZip.Count -gt 0) {
                Write-Host "The following items exist only in the ZIP file [$ZipFilePath]:"
                Write-Host ($OnlyInZip | Out-String)
            }
            $statusCode = -6501
        }
    }

    # Delete the temporary folder
    if ($statusCode -eq 0) {
        $statusCode = (Remove-Folder -FolderPath $TempExtractFolder)
    }
    
    return $statusCode
}

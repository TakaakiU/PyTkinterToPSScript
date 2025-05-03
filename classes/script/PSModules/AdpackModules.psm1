# Load required assemblies
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Create a folder to store the specified XML file
function New-XmlFolder {
    param (
        [System.String]$FilePath
    )

    [System.Int32]$statusCode = 0

    # Get the folder path to store the XML file
    $xmlDirectory = Split-Path -Path $FilePath
    # Create the folder if it does not exist
    if (-Not (Test-Path -Path $xmlDirectory)) {
        try {
            New-Item -ItemType Directory -Path $xmlDirectory | Out-Null
        }
        catch {
            Write-Error "An error occurred: $($_.Exception.Message)"
            $statusCode = -7101
        }
    }
}

# Function to calculate the SHA256 hash of a file
function Get-FileHashEX {
    param (
        [System.String]$FilePath,
        [ValidateSet("BASE64", "HEX")]
        [System.String]$HashFormat = "BASE64", # Specify "BASE64" or "HEX"
        [ValidateSet("SHA256", "SHA384", "SHA512", "SHA1", "MD5")]
        [System.String]$HashAlgorithm = "SHA256" # Hash algorithm to use
    )

    # Use Get-FileHash to retrieve the hash value
    $hashObject = Get-FileHash -Path $FilePath -Algorithm $HashAlgorithm

    # Convert and return the hash value based on the specified format
    switch ($HashFormat) {
        "BASE64" {
            # Convert HEX format to BASE64 format
            $hashBytes = [System.Convert]::FromHexString($hashObject.Hash)
            return [Convert]::ToBase64String($hashBytes)
        }
        "HEX" {
            # Return the HEX format as is
            return $hashObject.Hash
        }
        default {
            throw "Unsupported format: $HashFormat. Use 'BASE64' or 'HEX'."
        }
    }
}

# Function to generate an XML document
# Function to generate Index.xml
function New-IndexXml {
    param (
        [System.String]$FolderPath,  # Target folder to retrieve hash values
        [System.String]$Destination  # Save location for Index.xml
    )

    [System.Int32]$statusCode = 0

    # Get the current execution time
    $currentDate = (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz")

    # Retrieve the username and hostname
    $envUser = $env:USERNAME
    $envHost = $env:COMPUTERNAME

    try {
        # Create a new XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        $root = $xmlDoc.CreateElement("Index")
        $xmlDoc.AppendChild($root)

        # Create and add each element to the XML
        $title = $xmlDoc.CreateElement("Title")
        $title.InnerText = $FolderPath
        $root.AppendChild($title)

        $date = $xmlDoc.CreateElement("Date")
        $date.InnerText = $currentDate
        $root.AppendChild($date)

        $user = $xmlDoc.CreateElement("User")
        $user.InnerText = $envUser
        $root.AppendChild($user)

        $hostname = $xmlDoc.CreateElement("Host")
        $hostname.InnerText = $envHost
        $root.AppendChild($hostname)

        # Save the XML
        $xmlDoc.Save($Destination)
    }
    catch {
        Write-Error "An error occurred while creating META-INF/Index.xml. [$($_.Exception.Message)]"
        $statusCode = -7102
    }

    return $statusCode
}

function New-ManifestXml {
    param (
        [System.String]$FolderPath,
        [System.String]$Destination,
        [ValidateSet("BASE64", "HEX")]
        [System.String]$HashFormat = "BASE64", # Specify "BASE64" or "HEX"
        [ValidateSet("SHA256", "SHA384", "SHA512", "SHA1")]
        [System.String]$HashAlgorithm = "SHA256" # Hash algorithm to use
    )

    [System.Int32]$statusCode = 0

    try {
        # Create a new XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        
        # Create the root element with namespace
        $namespace = "http://www.w3.org/2000/09/xmldsig#"
        $root = $xmlDoc.CreateElement("Manifest", $namespace)
        $root.SetAttribute("Id", "bzpk-Manifest-0")
        $xmlDoc.AppendChild($root)
        
        # Calculate the hash value for each file and add to XML
        $files = Get-ChildItem -Recurse -Path $FolderPath | Where-Object {
            -Not $_.PSIsContainer -and $_.FullName -ne $Destination
        }
        foreach ($file in $files) {
            $relativePath = $file.FullName.Substring($FolderPath.Length + 1)
            $hash = Get-FileHashEX -FilePath $file.FullName

            $reference = $xmlDoc.CreateElement("Reference", $namespace)
            $reference.SetAttribute("URI", $relativePath)

            $digestMethod = $xmlDoc.CreateElement("DigestMethod", $namespace)
            # Change based on hash algorithm
            switch ($HashAlgorithm) {
                "SHA256" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256")
                }
                "SHA384" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmldsig-more#sha384")
                }
                "SHA512" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2001/04/xmlenc#sha512")
                }
                "SHA1" {
                    $digestMethod.SetAttribute("Algorithm", "http://www.w3.org/2000/09/xmldsig#sha1")
                }
            }
            $reference.AppendChild($digestMethod)

            $digestValue = $xmlDoc.CreateElement("DigestValue", $namespace)
            $digestValue.InnerText = $hash
            $reference.AppendChild($digestValue)

            $root.AppendChild($reference)
        }

        # Save the XML
        $xmlDoc.Save($Destination)
    }
    catch {
        Write-Error "An error occurred while creating META-INF/Manifest.xml. [$($_.Exception.Message)]"
        $statusCode = -7103
    }

    return $statusCode
}

function Test-HashValues {
    param (
        [System.String]$HashAlgorithm,
        [System.String]$ExtractedFolder
    )

    try {
        # Load the XML file
        $manifestFilePath = "$ExtractedFolder/META-INF/Manifest.xml"

        # Load the XML signature file
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($manifestFilePath)

        # Create and set the namespace manager
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager $xmlDoc.NameTable
        $namespaceManager.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")

        # Collect hash information within the signature
        $references = $xmlDoc.SelectNodes("//ds:Reference", $namespaceManager)

        $isVerified = $true

        foreach ($ref in $references) {
            $uri = $ref.GetAttribute("URI")
            $expectedHash = $ref.SelectSingleNode("ds:DigestValue", $namespaceManager).InnerText
            $digestAlgorithm = $ref.SelectSingleNode("ds:DigestMethod", $namespaceManager).GetAttribute("Algorithm")

            # Hash algorithm specified in the XML file
            switch ($digestAlgorithm) {
                "http://www.w3.org/2001/04/xmlenc#sha256" {
                    Write-Debug "SHA256 algorithm is used."
                }
                "http://www.w3.org/2001/04/xmlenc#sha512" {
                    Write-Debug "SHA512 algorithm is used."
                }
                "http://www.w3.org/2001/04/xmldsig-more#sha384" {
                    Write-Debug "SHA384 algorithm is used."
                }
                "http://www.w3.org/2000/09/xmldsig#sha1" {
                    Write-Debug "SHA1 algorithm is used."
                }
            }

            # Construct the file path
            $filePath = (Join-Path -Path $ExtractedFolder -ChildPath $uri)

            if (-Not (Test-Path -Path $filePath)) {
                Write-Host "File not found: $uri" -ForegroundColor Red
                $isVerified = $false
                continue
            }

            # Calculate and compare the hash
            $computedHash = (Get-FileHashEX -FilePath $filePath -Algorithm $HashAlgorithm)

            if ($computedHash -ne $expectedHash) {
                Write-Host "Hash mismatch: $uri" -ForegroundColor Red
                Write-Host "Expected hash: $expectedHash" -ForegroundColor Yellow
                Write-Host "Computed hash: $computedHash" -ForegroundColor Yellow
                $isVerified = $false
            }
        }

        if ($isVerified) {
            Write-Host "All files match the signature." -ForegroundColor Green
        } else {
            Write-Host "There are files that do not match." -ForegroundColor Red
            $statusCode = -7104
        }
    }
    catch {
        Write-Error "An error occurred: $_"
    }

    return $statusCode
}

function Compress-Package_Pack {
    param (
        [System.String]$ExePath,
        [System.String]$HashAlgorithm,
        [System.String]$FolderPath,
        [System.String]$ZipFilePath
    )

    [System.Int32]$statusCode = 0

    # PowerShell version check
    $statusCode = Test-PowerShellVersion7OrLater

    # Execute package generation
    if ($statusCode -eq 0) {
        try {
	        (& $ExePath -pack -hash $HashAlgorithm -in $FolderPath -out $ZipFilePath -force | Out-Null)
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7105
            }
        }
        catch {
	        Write-Error "An error occurred during package generation. [Target: $($FolderPath)]"
	        $statusCode = -7106
        }
    }

    # Compare the ZIP file and target folder after compression
    if ($statusCode -eq 0) {
        Write-Host "Step 3 Start: Compare before and after compression"
        Write-Host ""
        $statusCode = (Compare-ZipAndFolderContent -FolderPath $FolderPath -ZipFilePath $ZipFilePath)
    }

    return $statusCode
}

function Compress-Package_UserDefined {
    param (
        [System.String]$HashAlgorithm,
        [System.String]$FolderPath,
        [System.String]$ZipFilePath
    )

    # Specify the path for META-INF folder and XML files
    $indexFilePath = "$FolderPath/META-INF/Index.xml"
    $manifestFilePath = "$FolderPath/META-INF/Manifest.xml"

    [System.Int32]$statusCode = 0

    # PowerShell version check
    $statusCode = Test-PowerShellVersion7OrLater

    Write-Host "Step 1 Start: Create XML files"
    Write-Host ""
    # Create META/INF folder
    if ($statusCode -eq 0) {
        $statusCode = (New-DirectoryIfNotExists $xmlDirectory -ForceRecreate)
    }
    if ($statusCode -eq 0) {
        # Generate XML files
        $statusCode = (New-IndexXml -FolderPath $FolderPath -Destination $indexFilePath | Out-Null)
    }
    if ($statusCode -eq 0) {
        $statusCode = (New-ManifestXml -FolderPath $FolderPath -Destination $manifestFilePath -HashAlgorithm $HashAlgorithm | Out-Null)
    }

    # Compress the folder into a ZIP file
    if ($statusCode -eq 0) {
        Write-Host "Step 2 Start: Compress into ZIP file"
        Write-Host ""
        $statusCode = (Compress-FolderToZip -FolderPath $FolderPath -Destination $ZipFilePath)
    }

    # Compare the ZIP file and target folder after compression
    if ($statusCode -eq 0) {
        Write-Host "Step 3 Start: Compare before and after compression"
        Write-Host ""
        $statusCode = (Compare-ZipAndFolderContent -FolderPath $FolderPath -ZipFilePath $ZipFilePath)
        $resultData | ForEach-Object { Write-Host $_ }
    }

    return $statusCode
}

# Function to decompress a ZIP file
function Expand-Package_Adpack {
    param (
        [Switch]$Check,
        [Switch]$NoCheck,
        [System.String]$ExePath,
        [System.String]$HashAlgorithm,
        [System.String]$ZipFilePath,
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    # Check process
    if ($statusCode -eq 0) {
        if (-Not $NoCheck) {
            (& $ExePath -unpack -in $ZipFilePath -check | Out-Null)
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7107
            }
        }
        else {
            Write-Host "XML signature check skipped due to option specification."
        }
    }

    if ($statusCode -eq 0) {
        Write-Host "Unpack Command[$ExePath -unpack -in $ZipFilePath -out $FolderPath -force -nocheck | Out-Null]"
        Write-Host ""
        # Prepare output data if the option is not for checking
        if (-Not ($Check)) {
            (& $ExePath -unpack -in $ZipFilePath -out $FolderPath -force -nocheck | Out-Null)
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7108
            }
        }
        # Delete the temporary folder if the option is for checking only
        else {
            (& $ExePath -unpack -in $ZipFilePath -check | Out-Null)
            if ($LASTEXITCODE -ne 0) {
                $statusCode = -7109
            }
        }
    }
    
    return $statusCode
}

function Expand-Package_UserDefined {
    param (
        [Switch]$Check,
        [Switch]$NoCheck,
        [System.String]$HashAlgorithm,
        [System.String]$ZipFilePath,
        [System.String]$FolderPath
    )

    [System.Int32]$statusCode = 0

    # Create a temporary extraction folder
    if ($statusCode -eq 0) {
        $zipDirectory = (Split-Path -Path $ZipFilePath)
        $TempExtractFolder = Join-Path -Path ($zipDirectory) -ChildPath (".TempExtract_" + [System.Guid]::NewGuid())
        $statusCode = (New-DirectoryIfNotExists -FolderPath $TempExtractFolder)
    }

    # Decompress
    if ($statusCode -eq 0) {
        Write-Host "Step 1 Start: Decompress package data"
        Write-Host ""
        $statusCode = (Expand-ZipToTempFolder -ZipFilePath $ZipFilePath -TempFolderPath $TempExtractFolder)
    }

    # Compare META-INF/manifest.xml and actual data
    if ($statusCode -eq 0) {
        Write-Host "Step 2 Start: Compare manifest.xml and actual data"
        Write-Host ""
        if (-Not $NoCheck) {
            $statusCode = (Test-HashValues -HashAlgorithm $HashAlgorithm -ExtractedFolder $TempExtractFolder -ZipFilePath $ZipFilePath)
        }
        else {
            Write-Host "XML signature check skipped due to option specification."
        }
    }

    if ($statusCode -eq 0) {
        Write-Host "Step 3 Start: Post-processing"
        Write-Host ""
        # Prepare output data if the option is not for checking
        if (-Not ($Check)) {
            $statusCode = (Remove-Folder -FolderPath $FolderPath)

            if ($statusCode -eq 0) {
                try {
                    Move-Item -Path $TempExtractFolder -Destination $FolderPath

                }
                catch {
                    Write-Error "An error occurred while renaming the temporary folder [$($_.Exception.Message)]"
                    $statusCode = -7110
                }
            }
        }
        # Delete the temporary folder if the option is for checking only
        else {
            $statusCode = (Remove-Folder -FolderPath $TempExtractFolder)
        }
    }
    
    return $statusCode
}

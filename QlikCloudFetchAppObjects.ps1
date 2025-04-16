# Define input parameter for the script
param (
    [string]$masterItems = $false # Default value is $false, it can also be $true
)

Write-Host "Master items processing: $masterItems"
$currentDate = Get-Date -Format "yyyyMMdd"

# Define execution mode: "local" or "server"
$executionMode = "server" # Change to "server" for server execution

# Define paths based on execution mode
if ($executionMode -eq "server") {
    $basePath = "D:\Apps\QlikCLIBackup"
    $drivePath = "D:"
    # Set HTTPS proxy
    $env:HTTPS_PROXY = "http://proxy.company.net:8080"
} else {
    $basePath = "C:\Qlik-Cli\SaaS\companyScript\Application"
    $drivePath = "C:"
}

# Initialize execution
Set-Location -Path $drivePath
cd $drivePath

# Ensure the base path exists
if (-not (Test-Path $basePath)) {
    Write-Host "Base path does not exist: $basePath. Creating it..."
    New-Item -ItemType Directory -Path $basePath -Force | Out-Null
}

# Define paths
$logFile = "$basePath\Logs\ApplicationObjectsErrorLog.txt"
$logFilePath = "$basePath\Logs\ApplicationObjectsLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$destinationUrlTemplate = "https://<AzureStorage>.blob.core.windows.net/qlikcloud-backup/RepositorySnapshots/ApplicationObjects/Placeholder?<SAS TOKEN>"
if ($masterItems -eq $true) {
    $filePath = "$basePath\ApplicationObjectsMasterItems.csv"
    $AzureTargetFile = "$currentDate ApplicationObjectsMasterItems.csv"
    $destinationUrl = $destinationUrlTemplate -replace "Placeholder", $AzureTargetFile
} else {
    $filePath = "$basePath\ApplicationObjectsNoMasterItems.csv"
    $AzureTargetFile = "$currentDate ApplicationObjectsNoMasterItems.csv"
    $destinationUrl = $destinationUrlTemplate -replace "Placeholder", $AzureTargetFile
}

# Ensure the Logs directory exists
$logsPath = "$basePath\Logs"
if (-not (Test-Path $logsPath)) {
    Write-Host "Logs directory does not exist: $logsPath. Creating it..."
    New-Item -ItemType Directory -Path $logsPath -Force | Out-Null
}

# Initialize execution
Set-Location -Path $basePath

# Function to log errors
function Log-Error {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Initialize the CSV log file with headers
if (-not (Test-Path $logFilePath)) {
    "Timestamp,SpaceName,SpaceType,ApplicationId,ApplicationName,ObjectID,Status,Message" | Out-File -FilePath $logFilePath -Encoding UTF8
}

# Function to log to CSV
function Log-ToCsv {
    param (
        [string]$SpaceName,
        [string]$SpaceType,
        [string]$ApplicationId,
        [string]$ApplicationName,
        [string]$ObjectId,
        [string]$Status,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp,$SpaceName,$SpaceType,$ApplicationId,$ApplicationName,$ObjectId,$Status,$Message" | Out-File -FilePath $logFilePath -Append -Encoding UTF8
}

# Fetch space details
Write-Host "Fetching space details..."
$spaceCommand = "qlik space ls --limit 10000 --json"
try {
    $spaceJsonOutput = Invoke-Expression $spaceCommand | Out-String
    if ($spaceJsonOutput -match "Error:" -or $spaceJsonOutput -match "Usage:") {
        Write-Host "CLI Error fetching spaces: $spaceJsonOutput"
        Log-Error -Message "CLI Error fetching spaces: $spaceJsonOutput"
        exit 1
    }
    $spaces = $spaceJsonOutput | ConvertFrom-Json
} catch {
    Write-Host "Error executing space CLI command: $_"
    Log-Error -Message "Error executing space CLI command: $_"
    exit 1
}

# Create a hashtable for space mapping
$spaceMap = @{}
foreach ($space in $spaces) {
    $spaceMap[$space.id] = [PSCustomObject]@{
        SpaceName = $space.name
        SpaceType = $space.type
    }
}
Write-Host "Found $($spaceMap.Count) spaces."

# Fetch apps
Write-Host "Fetching apps..."
try {
    $apps = qlik app ls --limit 30000 --json | ConvertFrom-Json
    Write-Host "Number of apps fetched: $($apps.Count)"
} catch {
    Write-Host "Error fetching apps: $_"
    Log-Error -Message "Error fetching apps: $_"
    exit 1
}

# Filter for a specific app (for testing)
$testAppId = $null # Set to $null to process all apps
if ($testAppId) {
    $apps = $apps | Where-Object { $_.resourceId -eq $testAppId }
    Write-Host "Filtered to test App ID: $testAppId"
}

# Prepare output file
New-Item -Path $filePath -ItemType File -Force | Out-Null
$header = 'Space Name,Space Type,App ID,App Name,Object ID,Is Master Object,Object Type,Object Title'
Set-Content -Path $filePath -Value $header -Encoding UTF8

# Process apps
foreach ($app in $apps) {
    # Get space details for the app
    $spaceId = $app.resourceAttributes.spaceId
    $spaceInfo = if ($spaceId -and $spaceMap.ContainsKey($spaceId)) { $spaceMap[$spaceId] } else { $null }

    # Skip apps with empty or null SpaceType
    if (-not $spaceInfo -or -not $spaceInfo.SpaceType) {
        Write-Host "Skipping app $($app.resourceId) due to Personal space or missing space info."
        continue
    }

    # Skip apps that are not in managed spaces
    if ($spaceInfo.SpaceType -ne "managed") {
        Write-Host "Skipping app $($app.resourceId) in space type $($spaceInfo.SpaceType)"
        continue
    }

    # Skip spaces with specific naming conventions
    if ($spaceInfo.SpaceName -match "^(SELF DEV |RCD |INT )" -or $spaceInfo.SpaceName -match "^PRD.*(extract|pipeline)$") {
        Write-Host "Skipping app $($app.resourceId) in space $($spaceInfo.SpaceName) due to naming convention."
        continue
    }

    Write-Host "Processing app: $($app.name) ($($app.resourceId)) in space $($spaceInfo.SpaceName) ($($spaceInfo.SpaceType))"
    $metadata = @()
    try {
        $tempMetadata = qlik app object ls -a $app.resourceId --no-data --json | ConvertFrom-Json
        Write-Host "Number of objects fetched: $($tempMetadata.Count)"
    } catch {
        Write-Host "Error fetching objects for app $($app.resourceId): $_"
        Log-ToCsv -spaceName $spaceInfo.SpaceName -spaceType $spaceInfo.SpaceType -applicationId $app.resourceId -applicationName $app.name -objectId "" -status "Failure" -message "Error fetching objects: $_"
        continue
    }

    # Filter for a specific app object (for testing)
    $testAppObjectId = $null # Set to "05ea2024-717a-430e-a410-7034bef0fa2d" or $null to process all objects
    if ($testAppObjectId) {
        $tempMetadata = $tempMetadata | Where-Object { $_.qId -eq $testAppObjectId }
        Write-Host "Filtered to test App Object ID: $testAppObjectId"
    }

    $maxRetries = 3 # Maximum number of retries for each object
    $retryDelay = 5 # Seconds to wait between retries

    foreach ($object in $tempMetadata) {
        Write-Host "Object type: $($object.qType)"
        $retryCount = 0
        $success = $false

        if ($object.qType -eq "masterobject") {
            if ($masterItems -eq $true) {
                Write-Host "Process masterobject (fetch visualization and title)"
                while ($retryCount -lt $maxRetries -and -not $success) {
                    try {
                        Write-Host "Captured masterobject with qType: $($object.qType) for object ID: $($object.qId)"
                        # Validate app and object IDs
                        if (-not $app.resourceId -or -not $object.qId) {
                            throw "Missing App ID or Object ID. App ID: $($app.resourceId), Object ID: $($object.qId)"
                        }

                        # Capture raw JSON output as a string
                        $objectPropertiesJson = qlik app object properties -a $app.resourceId $object.qId --no-data --json
                        if ([string]::IsNullOrWhiteSpace($objectPropertiesJson)) {
                            throw "Could not retrieve properties for object ID '$($object.qId)'. JSON is empty or null."
                        }
                        
                        $jsonString = $objectPropertiesJson -join ""

                        # Find qMetaDef
                        $qMetaDefIndex = $jsonString.IndexOf('"qMetaDef":')
                        if ($qMetaDefIndex -ge 0) {
                            $qMetaDefSubstring = $jsonString.Substring($qMetaDefIndex, [Math]::Min(200, $jsonString.Length - $qMetaDefIndex))
                            if ($qMetaDefSubstring -match '"title"\s*:\s*"([^"]+)"') {
                                $title = $Matches[1].Trim() # Trim leading/trailing spaces
                            } else {
                                $title = ""
                            }
                        } else {
                            $title = ""
                        }

                        # Find visualization
                        $visualizationIndex = $jsonString.IndexOf('"visualization":')
                        if ($visualizationIndex -ge 0) {
                            $visualizationSubstring = $jsonString.Substring($visualizationIndex, [Math]::Min(200, $jsonString.Length - $visualizationIndex))
                            if ($visualizationSubstring -match '"visualization"\s*:\s*"([^"]+)"') {
                                $visualization = $Matches[1] # Keep internal spaces, remove quotes
                            } else {
                                $visualization = ""
                            }
                        } else {
                            $visualization = ""
                        }

                        # Check if visualization property exists and is not null
                        if ($visualization -ne $null) {
                            $objectMetadata = [PSCustomObject]@{
                                "Space Name"      = $spaceInfo.SpaceName
                                "Space Type"      = $spaceInfo.SpaceType
                                "App ID"          = $app.resourceId
                                "App Name"        = $app.name
                                "Object ID"       = $object.qId
                                "Is Master Object" = 1
                                "Object Type"     = $visualization.ToString()
                                "Object Title"    = ($title -replace '[\r\n]', '')
                            }
                            $metadata += $objectMetadata
                            Write-Host "Captured visualization: $visualization for object ID: $($object.qId)"
                        } else {
                            Write-Host "Skipping object ID: $($object.qId) as it has no visualization property"
                        }

                        $success = $true
                        Log-ToCsv -spaceName $spaceInfo.SpaceName -spaceType $spaceInfo.SpaceType -applicationId $app.resourceId -applicationName $app.name -objectId $object.qId -status "Success" -message "Masterobject processed successfully."
                    } catch {
                        $retryCount++
                        $errorMessage = "Error processing object ID '$($object.qId)' (Attempt $retryCount/$maxRetries): $_"
                        Write-Host $errorMessage
                        Log-ToCsv -spaceName $spaceInfo.SpaceName -spaceType $spaceInfo.SpaceType -applicationId $app.resourceId -applicationName $app.name -objectId $object.qId -status "Failure" -message $errorMessage
                        if ($retryCount -ge $maxRetries) {
                            Write-Host "Max retries reached for object ID '$($object.qId)'. Skipping."
                        } else {
                            Start-Sleep -Seconds $retryDelay
                        }
                    }
                }
            } else {
                # Log masterobject existence without processing
                Write-Host "Logging masterobject existence for object ID: $($object.qId) (masterItems = $false)"
                $objectMetadata = [PSCustomObject]@{
                    "Space Name"      = $spaceInfo.SpaceName
                    "Space Type"      = $spaceInfo.SpaceType
                    "App ID"          = $app.resourceId
                    "App Name"        = $app.name
                    "Object ID"       = $object.qId
                    "Is Master Object" = 1
                    "Object Type"     = ""
                    "Object Title"    = ""
                }
                $metadata += $objectMetadata
                Log-ToCsv -spaceName $spaceInfo.SpaceName -spaceType $spaceInfo.SpaceType -applicationId $app.resourceId -applicationName $app.name -objectId $object.qId -status "Success" -message "Masterobject logged without processing (masterItems = $false)."
            }
        } else {
            # Handle non-masterobject
            $objectMetadata = [PSCustomObject]@{
                "Space Name"      = $spaceInfo.SpaceName
                "Space Type"      = $spaceInfo.SpaceType
                "App ID"          = $app.resourceId
                "App Name"        = $app.name
                "Object ID"       = $object.qId
                "Is Master Object" = 0
                "Object Type"     = $object.qType
                "Object Title"    = ($object.title -replace '[\r\n]', '')
            }
            $metadata += $objectMetadata
            Write-Host "Captured non-masterobject with qType: $($object.qType) for object ID: $($object.qId)"
            Log-ToCsv -spaceName $spaceInfo.SpaceName -spaceType $spaceInfo.SpaceType -applicationId $app.resourceId -applicationName $app.name -objectId $object.qId -status "Success" -message "Non-masterobject processed successfully."
        }
    }

    if ($metadata.Count -gt 0) {
        $metadata | Export-Csv -Path $filePath -Append -NoTypeInformation -Encoding UTF8
    }
}

Write-Host "Script execution completed. Output written to $filePath"
Write-Host "Master items processed: $masterItems"

# Upload the CSV file
azcopy copy "$filePath" "$destinationUrl"

# CBZ Comic File Monitor and Tagger Script
# Monitors a directory for new CBZ files, tags them with ComicTagger, modifies metadata, and renames them

# Configuration - Built-in parameters
$WatchDirectory = "C:\path\to\comic\directory"
$ComicTaggerPath = "C:\path\to\comictagger.exe"
$LogFilePath = "C:\path\to\log\file"
$RetryInterval = 1200  # 5 minutes in seconds

# Global variables
$script:RetryQueue = @()
$script:ProcessedFiles = @{}

# Function to log messages with timestamp
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console
    Write-Host $logMessage
    
    # Write to log file
    try {
        # Ensure log directory exists
        $logDir = Split-Path $LogFilePath -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        # Append to log file
        Add-Content -Path $LogFilePath -Value $logMessage -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write to log file: $($_.Exception.Message)"
    }
}

# Function to test if ComicTagger is available
function Test-ComicTagger {
    try {
        $result = & $ComicTaggerPath --version 2>$null
        return $true
    }
    catch {
        Write-Log "ComicTagger not found at: $ComicTaggerPath" "ERROR"
        return $false
    }
}

# Function to modify ComicInfo.xml to combine Series and Volume
function Update-ComicInfoXml {
    param([string]$CbzPath)
    
    try {
        Write-Log "Updating ComicInfo.xml for: $CbzPath"
        
        # Create a temporary directory for extraction
        $tempDir = Join-Path $env:TEMP "CBZ_Temp_$(Get-Random)"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        try {
            # Load required assembly for ZIP operations
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            
            # Extract the CBZ file
            [System.IO.Compression.ZipFile]::ExtractToDirectory($CbzPath, $tempDir)
            
            # Look for ComicInfo.xml
            $comicInfoPath = Join-Path $tempDir "ComicInfo.xml"
            
            if (-not (Test-Path $comicInfoPath)) {
                Write-Log "ComicInfo.xml not found in: $CbzPath" "WARNING"
                return $false
            }
            
            # Load the XML file
            [xml]$comicInfo = Get-Content $comicInfoPath
            
            # Get the current series and volume values
            $currentSeries = $comicInfo.ComicInfo.Series
            $currentVolume = $comicInfo.ComicInfo.Volume
            
            Write-Log "Current Series: '$currentSeries', Volume: '$currentVolume'" "INFO"
            
            # Check if both series and volume exist and are not empty
            if ([string]::IsNullOrWhiteSpace($currentSeries)) {
                Write-Log "Series tag is empty or missing in: $CbzPath" "WARNING"
                return $false
            }
            
            if ([string]::IsNullOrWhiteSpace($currentVolume)) {
                Write-Log "Volume tag is empty or missing in: $CbzPath" "INFO"
                return $false
            }
            
            # Create the new series value in the format "Series Vol. X"
            $newSeries = "$currentSeries Vol. $currentVolume"
            Write-Log "Updating series to: '$newSeries'" "INFO"
            
            # Update the series tag
            $comicInfo.ComicInfo.Series = $newSeries
            
            # Save the modified XML back to the file
            $comicInfo.Save($comicInfoPath)
            Write-Log "Updated ComicInfo.xml with new series value" "SUCCESS"
            
            # Create a new CBZ file with the updated ComicInfo.xml
            Remove-Item $CbzPath -Force
            [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $CbzPath)
            
            Write-Log "Successfully recreated CBZ file with updated metadata" "SUCCESS"
            return $true
        }
        finally {
            # Clean up temporary directory
            if (Test-Path $tempDir) {
                Remove-Item $tempDir -Recurse -Force
            }
        }
    }
    catch {
        Write-Log "Error updating ComicInfo.xml for $CbzPath : $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to get metadata from CBZ file using ComicTagger
function Get-ComicMetadata {
    param([string]$FilePath)
    
    try {
        Write-Log "Attempting to tag file: $FilePath"
        
        # First, try to tag the file with ComicTagger (both ComicRack and ComicBookInfo formats)
        $tagResult = & $ComicTaggerPath --online -s -f --tags-write "CR,CIX" "$FilePath" 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            Write-Log "ComicTagger failed to tag file: $FilePath" "WARNING"
            return $null
        }
        
        Write-Log "Successfully tagged file with ComicTagger: $FilePath" "SUCCESS"
        
        # Update ComicInfo.xml to combine series and volume
        $updateResult = Update-ComicInfoXml -CbzPath $FilePath
        if ($updateResult) {
            Write-Log "Successfully updated ComicInfo.xml metadata" "SUCCESS"
        }
        
        # Now extract the metadata using print with JSON format
        $metadataOutput = & $ComicTaggerPath --print --json --tags-read "CR" "$FilePath" 2>&1
        
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Failed to extract metadata from: $FilePath" "ERROR"
            return $null
        }
        
        # Parse the mixed text/JSON output - find the JSON part
        try {
            $outputText = $metadataOutput -join "`n"
            
            # The JSON starts after the text output, look for the opening brace
            $jsonStart = $outputText.IndexOf('{')
            
            if ($jsonStart -ge 0) {
                # Extract everything from the first { to the end
                $jsonText = $outputText.Substring($jsonStart)
                $jsonData = $jsonText | ConvertFrom-Json
                
                $metadata = @{
                    'Series' = $jsonData.md.series
                    'Volume' = $jsonData.md.volume
                    'Issue' = $jsonData.md.issue
                    'Year' = $jsonData.md.year
                }
            }
            else {
                Write-Log "No JSON data found in ComicTagger output for: $FilePath" "ERROR"
                return $null
            }
        }
        catch {
            Write-Log "Failed to parse JSON metadata from: $FilePath - $($_.Exception.Message)" "ERROR"
            Write-Log "ComicTagger output length: $($outputText.Length) characters" "DEBUG"
            return $null
        }
        
        # Log the extracted metadata
        Write-Log "Extracted metadata for: $FilePath" "INFO"
        $series = if ($metadata.ContainsKey('Series') -and $metadata['Series']) { $metadata['Series'] } else { 'Not found' }
        $volume = if ($metadata.ContainsKey('Volume') -and $metadata['Volume']) { $metadata['Volume'] } else { 'Not found' }
        $issue = if ($metadata.ContainsKey('Issue') -and $metadata['Issue']) { $metadata['Issue'] } else { 'Not found' }
        $year = if ($metadata.ContainsKey('Year') -and $metadata['Year']) { $metadata['Year'] } else { 'Not found' }
        
        Write-Log "  Series: $series" "INFO"
        Write-Log "  Volume: $volume" "INFO"
        Write-Log "  Issue: $issue" "INFO"
        Write-Log "  Year: $year" "INFO"
        
        return $metadata
        
    }
    catch {
        Write-Log "Error processing file $FilePath : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Function to generate new filename from metadata
function Get-NewFileName {
    param(
        [hashtable]$Metadata,
        [string]$OriginalExtension
    )
    
    $series = $Metadata['Series'] -replace '[<>:"/\\|?*]', '_'
    $volume = $Metadata['Volume']
    $issue = $Metadata['Issue']
    $year = $Metadata['Year']
    
    if (-not $series) {
        Write-Log "No series found in metadata, cannot rename file" "WARNING"
        return $null
    }
    
    # Build the filename template
    # Note: After our ComicInfo.xml update, the series should already include volume info
    # So we'll use the series as-is and only add volume if it's not already included
    $newName = $series
    
    # Only add volume if it's not already part of the series name and exists
    if ($volume -and $volume -ne "" -and $volume -ne "None" -and $series -notmatch "Vol\.\s*\d+") {
        $newName += " Vol.$volume"
    }
    
    if ($issue) {
        # Pad issue number to 3 digits
        if ($issue -match '^\d+$') {
            $paddedIssue = $issue.PadLeft(3, '0')
        } else {
            $paddedIssue = $issue
        }
        $newName += " #$paddedIssue"
    }
    
    if ($year) {
        $newName += " ($year)"
    }
    
    # Clean up the filename and add extension
    $newName = $newName -replace '[<>:"/\\|?*]', '_'
    $newName = $newName + $OriginalExtension
    
    return $newName
}

# Function to process a CBZ file
function Process-CBZFile {
    param([string]$FilePath)
    
    Write-Log "Processing file: $FilePath"
    
    # Check if file still exists
    if (-not (Test-Path $FilePath)) {
        Write-Log "File no longer exists: $FilePath" "WARNING"
        return $true  # Consider it processed to remove from queue
    }
    
    # Wait a moment to ensure file is not being written to
    Start-Sleep -Seconds 2
    
    # Get metadata (this will also update the ComicInfo.xml)
    $metadata = Get-ComicMetadata -FilePath $FilePath
    
    if (-not $metadata) {
        Write-Log "Failed to get metadata for: $FilePath" "WARNING"
        return $false  # Failed, should retry
    }
    
    # Generate new filename
    $directory = Split-Path $FilePath -Parent
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $newFileName = Get-NewFileName -Metadata $metadata -OriginalExtension $extension
    
    if (-not $newFileName) {
        Write-Log "Could not generate new filename for: $FilePath" "WARNING"
        return $true  # Consider processed even if we can't rename
    }
    
    $newFilePath = Join-Path $directory $newFileName
    
    # Check if target filename already exists
    if (Test-Path $newFilePath) {
        $counter = 1
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($newFileName)
        do {
            $newFileName = "$baseName ($counter)$extension"
            $newFilePath = Join-Path $directory $newFileName
            $counter++
        } while (Test-Path $newFilePath)
    }
    
    # Rename the file
    try {
        Rename-Item -Path $FilePath -NewName $newFileName -ErrorAction Stop
        Write-Log "Successfully renamed '$FilePath' to '$newFileName'" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to rename file: $($_.Exception.Message)" "ERROR"
        return $true  # Consider processed to avoid infinite retries for permission issues
    }
}

# Function to process retry queue
function Process-RetryQueue {
    Write-Log "Processing retry queue. Items in queue: $($script:RetryQueue.Count)"
    
    $itemsToRemove = @()
    
    for ($i = 0; $i -lt $script:RetryQueue.Count; $i++) {
        $item = $script:RetryQueue[$i]
        
        # Check if enough time has passed since last attempt
        if ((Get-Date) -ge $item.NextAttempt) {
            Write-Log "Retrying file: $($item.FilePath)"
            
            if (Process-CBZFile -FilePath $item.FilePath) {
                # Successfully processed, remove from queue
                $itemsToRemove += $i
            }
            else {
                # Update next attempt time
                $script:RetryQueue[$i].NextAttempt = (Get-Date).AddSeconds($RetryInterval)
                $script:RetryQueue[$i].AttemptCount++
                Write-Log "Retry failed for: $($item.FilePath). Attempt: $($script:RetryQueue[$i].AttemptCount)"
            }
        }
    }
    
    # Remove successfully processed items (in reverse order to maintain indices)
    $itemsToRemove | Sort-Object -Descending | ForEach-Object {
        $script:RetryQueue.RemoveAt($_)
    }
}

# Function to add file to retry queue
function Add-ToRetryQueue {
    param([string]$FilePath)
    
    $retryItem = @{
        FilePath = $FilePath
        NextAttempt = (Get-Date).AddSeconds($RetryInterval)
        AttemptCount = 1
    }
    
    $script:RetryQueue += $retryItem
    Write-Log "Added to retry queue: $FilePath"
}

# Function to handle new file events
function Handle-NewFile {
    param([string]$FilePath)
    
    # Skip if already processed or in queue
    if ($script:ProcessedFiles.ContainsKey($FilePath)) {
        return
    }
    
    # Check if already in retry queue
    if ($script:RetryQueue | Where-Object { $_.FilePath -eq $FilePath }) {
        return
    }
    
    Write-Log "New CBZ file detected: $FilePath"
    $script:ProcessedFiles[$FilePath] = $true
    
    if (-not (Process-CBZFile -FilePath $FilePath)) {
        Add-ToRetryQueue -FilePath $FilePath
    }
}

# Main monitoring function
function Start-CBZMonitoring {
    Write-Log "=== CBZ Comic File Monitor Started ===" "INFO"
    Write-Log "Configuration:" "INFO"
    Write-Log "  Watch Directory: $WatchDirectory" "INFO"
    Write-Log "  ComicTagger Path: $ComicTaggerPath" "INFO"
    Write-Log "  Log File: $LogFilePath" "INFO"
    Write-Log "  Retry Interval: $RetryInterval seconds" "INFO"
    Write-Log "=================================" "INFO"
    
    # Validate directory exists
    if (-not (Test-Path $WatchDirectory)) {
        Write-Log "Watch directory does not exist: $WatchDirectory" "ERROR"
        return
    }
    
    # Validate ComicTagger is available
    if (-not (Test-ComicTagger)) {
        return
    }
    
    Write-Log "Monitoring for NEW CBZ files only (existing files will be ignored)"
    Write-Log "Files will be tagged, have ComicInfo.xml updated (Series + Volume combined), and renamed"
    
    # Set up file system watcher
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $WatchDirectory
    $watcher.Filter = "*.cbz"
    $watcher.EnableRaisingEvents = $true
    
    # Register event handler
    $action = {
        $path = $Event.SourceEventArgs.FullPath
        $name = $Event.SourceEventArgs.Name
        $changeType = $Event.SourceEventArgs.ChangeType
        
        if ($changeType -eq 'Created') {
            Handle-NewFile -FilePath $path
        }
    }
    
    Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $action
    
    Write-Log "File system watcher started. Press Ctrl+C to stop monitoring."
    
    # Main monitoring loop
    try {
        while ($true) {
            Start-Sleep -Seconds 30
            
            # Process retry queue
            if ($script:RetryQueue.Count -gt 0) {
                Process-RetryQueue
            }
            
            # Clean up processed files list periodically (keep last 1000 entries)
            if ($script:ProcessedFiles.Count -gt 1000) {
                $keysToRemove = ($script:ProcessedFiles.Keys | Select-Object -First ($script:ProcessedFiles.Count - 1000))
                $keysToRemove | ForEach-Object { $script:ProcessedFiles.Remove($_) }
            }
        }
    }
    catch {
        Write-Log "Monitoring stopped: $($_.Exception.Message)" "INFO"
    }
    finally {
        # Cleanup
        $watcher.EnableRaisingEvents = $false
        $watcher.Dispose()
        Write-Log "File system watcher stopped."
    }
}

# Start the monitoring
Start-CBZMonitoring


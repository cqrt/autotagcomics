# Comic Auto-Tagger and Renamer Script for Windows
# Monitors a directory, tags new comics with metadata, and renames them
# If a file is not tagged, rename with [untagged] and retry once per hour
# Renaming template: {series} Vol.{volume} #{issue} ({year}) - Excludes volume if missing

# === CONFIGURATION SECTION ===
$watchFolder = "C:\Users\lemck\Comics\Auto Add"  # Directory to monitor
$comictagger = "C:\Comictagger\comictagger.exe"  # Path to Comictagger
$logFile = "C:\Comictagger\Logs\Comictagger.log"  # Log file (create folder first)
$retryInterval = 3600  # Seconds between retries (5 minutes)

# === FILE SYSTEM WATCHER SETUP ===
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchFolder
$watcher.Filter = "*.cbz"
$watcher.EnableRaisingEvents = $true
$watcher.IncludeSubdirectories = $false

# === PROCESS NEW FILE FUNCTION ===
function Process-NewComicFile {
    param($filePath)
    
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $directory = [System.IO.Path]::GetDirectoryName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath)
    
    # Wait briefly to ensure file is fully written
    Start-Sleep -Seconds 5
    
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Processing: $fileName" | Out-File $logFile -Append
    
    try {
        # Attempt tagging
        & $comictagger --online -s -f --tags-write "CR,CIX" $filePath
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Tagged: $fileName" | Out-File $logFile -Append
        
        # Attempt metadata extraction
        $rawOutput = & $comictagger -p --json $filePath 2>&1 | Out-String
        
        # Extract JSON from mixed output
        $jsonStart = $rawOutput.IndexOf('{')
        $jsonEnd = $rawOutput.LastIndexOf('}')
        
        if ($jsonStart -ge 0 -and $jsonEnd -gt $jsonStart) {
            $jsonPart = $rawOutput.Substring($jsonStart, $jsonEnd - $jsonStart + 1)
            $metadata = $jsonPart | ConvertFrom-Json -ErrorAction Stop
            
            # === METADATA EXTRACTION ===
            $tagData = $null
            
            # Structure 1: Metadata directly in "md" property
            if ($metadata.PSObject.Properties.Name -contains 'md') {
                $tagData = $metadata.md
            }
            # Structure 2: Metadata in first array element
            elseif ($metadata -is [array] -and $metadata.Count -gt 0 -and $metadata[0].PSObject.Properties.Name -contains 'metadata') {
                $tagData = $metadata[0].metadata
            }
            # Structure 3: Metadata at root level
            elseif ($metadata.PSObject.Properties.Name -contains 'series') {
                $tagData = $metadata
            }
            else {
                throw "Unrecognized metadata structure"
            }
            
            # === SERIES NAME ===
            $series = $null
            $propertyNames = @('series', 'Series', 'SERIES')
            foreach ($prop in $propertyNames) {
                if ($tagData.PSObject.Properties.Name -contains $prop -and 
                    -not [string]::IsNullOrWhiteSpace($tagData.$prop)) {
                    $series = $tagData.$prop
                    break
                }
            }
            
            if ([string]::IsNullOrWhiteSpace($series)) {
                throw "Series name missing in metadata"
            }
            
            # Sanitize series name
            $series = $series -replace '[\\/:*?"<>|]', ''
            $series = $series.Trim()
            
            # === VOLUME HANDLING ===
            $volumePart = ""
            $volumeValue = $null
            $volumeProps = @('volume', 'Volume', 'VOLUME')
            foreach ($prop in $volumeProps) {
                if ($tagData.PSObject.Properties.Name -contains $prop) {
                    $volumeValue = $tagData.$prop
                    break
                }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($volumeValue)) {
                $volumePart = " Vol.$volumeValue"
            }
            
            # === ISSUE NUMBER ===
            $issue = $null
            $issueProps = @('issue', 'Issue', 'ISSUE', 'number', 'Number')
            foreach ($prop in $issueProps) {
                if ($tagData.PSObject.Properties.Name -contains $prop -and 
                    -not [string]::IsNullOrWhiteSpace($tagData.$prop)) {
                    $issue = $tagData.$prop
                    break
                }
            }
            
            if ([string]::IsNullOrWhiteSpace($issue)) {
                throw "Issue number missing in metadata"
            }
            
            # Clean issue number
            $issue = $issue.Trim()
            $issue = $issue -replace '^0+(\d+\.?\d*)', '$1'
            
            # === PUBLICATION YEAR ===
            $year = $null
            $yearProps = @('year', 'Year', 'YEAR', 'coverYear', 'publicationYear')
            foreach ($prop in $yearProps) {
                if ($tagData.PSObject.Properties.Name -contains $prop -and 
                    -not [string]::IsNullOrWhiteSpace($tagData.$prop)) {
                    $year = $tagData.$prop
                    break
                }
            }
            
            if ([string]::IsNullOrWhiteSpace($year)) {
                $year = (Get-Date).Year
            }
            
            # === BUILD NEW FILENAME ===
            $newName = "${series}${volumePart} #${issue} (${year})${extension}"
            $newPath = Join-Path -Path $directory -ChildPath $newName
            
            # === RENAME FILE ===
            Rename-Item -Path $filePath -NewName $newName -Force
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Renamed: $fileName -> $newName" | Out-File $logFile -Append
        }
        else {
            throw "No valid JSON found in comictagger output"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | PROCESS ERROR: $fileName - $errorMsg" | Out-File $logFile -Append
        
        # Mark file as untagged
        $newName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + " [untagged]" + $extension
        $newPath = Join-Path -Path $directory -ChildPath $newName
        Rename-Item -Path $filePath -NewName $newName -Force
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Marked as untagged: $newName" | Out-File $logFile -Append
    }
}

# === RETRY UNTAGGED FILES FUNCTION ===
function Retry-UntaggedFiles {
    $untaggedFiles = Get-ChildItem -Path $watchFolder -Filter "* [untagged].cbz" -File
    
    foreach ($file in $untaggedFiles) {
        $filePath = $file.FullName
        $fileName = $file.Name
        
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Retrying: $fileName" | Out-File $logFile -Append
        
        try {
            # Remove [untagged] marker temporarily
            $tempName = $fileName -replace " \[untagged\]", ""
            $tempPath = Join-Path -Path $watchFolder -ChildPath $tempName
            Rename-Item -Path $filePath -NewName $tempName -Force
            
            # Process the file normally
            Process-NewComicFile -filePath $tempPath
        }
        catch {
            $errorMsg = $_.Exception.Message
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | RETRY FAILED: $fileName - $errorMsg" | Out-File $logFile -Append
            
            # Re-add [untagged] marker if still failed
            if (Test-Path $tempPath) {
                $markedName = [System.IO.Path]::GetFileNameWithoutExtension($tempName) + " [untagged]" + $extension
                Rename-Item -Path $tempPath -NewName $markedName -Force
            }
        }
    }
}

# === EVENT HANDLER FOR NEW FILES ===
$newFileAction = {
    $filePath = $Event.SourceEventArgs.FullPath
    $fileName = $Event.SourceEventArgs.Name
    
    # Skip [untagged] files
    if ($fileName -match "\[untagged\]") {
        return
    }
    
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | New file detected: $fileName" | Out-File $logFile -Append
    Process-NewComicFile -filePath $filePath
}

# Register events
$eventHandler = Register-ObjectEvent -InputObject $watcher -EventName "Created" -Action $newFileAction

# === MAIN LOOP FOR RETRYING UNTAGGED FILES ===
"$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Script started" | Out-File $logFile -Append

try {
    while ($true) {
        # Retry untagged files
        Retry-UntaggedFiles
        
        # Log status
        $untaggedCount = (Get-ChildItem -Path $watchFolder -Filter "* [untagged].cbz" -File).Count
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Status: $untaggedCount untagged files pending" | Out-File $logFile -Append
        
        # Wait for next cycle
        Start-Sleep -Seconds $retryInterval
    }
}
finally {
    # Cleanup on exit
    $eventHandler | Unregister-Event
    $watcher.Dispose()
}

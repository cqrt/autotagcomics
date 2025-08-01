# Comic Auto-Tagger and Renamer Script for Windows
# Monitors a directory, tags new comics with metadata, and renames them
# Renaming template: {series} Vol.{volume} #{issue} ({year}) - Excludes volume if missing

# === CONFIGURATION SECTION ===
$watchFolder = "C:\Users\lemck\Comics\Auto Add"  # Directory to monitor
$comictagger = "C:\Comictagger\comictagger.exe"  # Path to Comictagger
$logFile = "C:\Comictagger\Logs\Comictagger.log"  # Log file (create folder first)

# === FILE SYSTEM WATCHER SETUP ===
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchFolder
$watcher.Filter = "*.cbz"
$watcher.EnableRaisingEvents = $true

# === EVENT HANDLER FOR NEW FILES ===
Register-ObjectEvent $watcher "Created" -Action {
    $filePath = $Event.SourceEventArgs.FullPath
    $fileName = $Event.SourceEventArgs.Name
    $directory = [System.IO.Path]::GetDirectoryName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath)
    
    Start-Sleep -Seconds 10
    
    # === STEP 1: TAG THE COMIC ===
    try {
        & $comictagger --online -s -f --tags-write "CR,CIX" $filePath
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Tagged: $fileName" | Out-File $logFile -Append
    }
    catch {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | TAG ERROR: $fileName - $_" | Out-File $logFile -Append
        return
    }
    
    # === STEP 2: EXTRACT METADATA ===
    try {
        # Capture raw output (mixed text and JSON)
        $rawOutput = & $comictagger -p --json $filePath 2>&1 | Out-String
        
        # Extract JSON from mixed output
        $jsonStart = $rawOutput.IndexOf('{')
        $jsonEnd = $rawOutput.LastIndexOf('}')
        
        if ($jsonStart -ge 0 -and $jsonEnd -gt $jsonStart) {
            $jsonPart = $rawOutput.Substring($jsonStart, $jsonEnd - $jsonStart + 1)
            $metadata = $jsonPart | ConvertFrom-Json -ErrorAction Stop
            
            # Process metadata
            $tagData = $metadata.md
            
            # === SERIES NAME ===
            $series = if (-not [string]::IsNullOrWhiteSpace($tagData.series)) { 
                ($tagData.series -replace '[\\/:*?"<>|]', '').Trim()
            } else { 
                throw "Series name missing in metadata" 
            }
            
            # === VOLUME HANDLING ===
            $volumePart = ""
            $volumeValue = if ($tagData.volume -is [string]) {
                $tagData.volume
            } elseif ($tagData.volume -is [int]) {
                $tagData.volume.ToString()
            } else {
                $null
            }
            
            if (-not [string]::IsNullOrWhiteSpace($volumeValue)) {
                $volumePart = " Vol.$volumeValue"
            }
            
            # === ISSUE NUMBER ===
            $issue = if (-not [string]::IsNullOrWhiteSpace($tagData.issue)) { 
                $cleanIssue = $tagData.issue.Trim()
                $cleanIssue -replace '^0+(\d+\.?\d*)', '$1'
            } else { 
                throw "Issue number missing in metadata"
            }
            
            # === PUBLICATION YEAR ===
            $year = if ($tagData.year -is [int]) { 
                $tagData.year.ToString()
            } elseif (-not [string]::IsNullOrWhiteSpace($tagData.year)) {
                $tagData.year
            } else { 
                (Get-Date).Year.ToString()
            }
            
            # === STEP 3: BUILD NEW FILENAME ===
            $newName = "${series}${volumePart} #${issue} (${year})${extension}"
            $newPath = Join-Path -Path $directory -ChildPath $newName
            
            # DEBUG: Log filename components
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | Filename components:`nSeries: '$series'`nVolume: '$volumeValue'`nIssue: '$issue'`nYear: '$year'`nNew Name: '$newName'" | Out-File $logFile -Append
            
            # === STEP 4: RENAME FILE ===
            if (Test-Path $newPath) {
                "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | SKIPPED: $fileName (Target exists)" | Out-File $logFile -Append
            }
            else {
                Rename-Item -Path $filePath -NewName $newName -Force
                "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | RENAMED: $fileName -> $newName" | Out-File $logFile -Append
            }
        }
        else {
            throw "No valid JSON found in comictagger output"
        }
    }
    catch {
        $errorMsg = $_.Exception.Message
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | METADATA ERROR: $fileName - $errorMsg" | Out-File $logFile -Append
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm') | SKIPPED RENAMING: $fileName (Metadata extraction failed)" | Out-File $logFile -Append
        
        # NOTICE: File is left with original name
    }
}

# === MAIN LOOP ===
while ($true) { 
    Start-Sleep -Seconds 60
}
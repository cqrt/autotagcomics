## Comic Auto-Tagger and Renamer Script for Windows

- Monitors a directory, tags new comics with metadata, and renames them
- If metadata is not found, it will queue the file and retry every hour
- Renaming template: {series} Vol.{volume} #{issue} ({year}) - Excludes volume if missing from metadata
- Pads issue number to 3 digits
- If Volume exists in metadata, combines Series and Volume and writes back to ComicInfo.xml Series tag

#### CONFIGURATION

- $WatchDirectory = "C:\path\to\comics"  # Directory to monitor
- $ComicTaggerPath = "C:\path\to\comictagger.exe"  # Path to Comictagger
- $LogFilePath = "C:\path\to\logs\Comictagger.log"  # Log file (create folder first)
- $RetryInterval = 3600 # in seconds

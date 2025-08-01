Comic Auto-Tagger and Renamer Script for Windows
Monitors a directory, tags new comics with metadata, and renames them
Renaming template: {series} Vol.{volume} #{issue} ({year}) - Excludes volume if missing

CONFIGURATION
$watchFolder = "/path/to/comics"  # Directory to monitor
$comictagger = "/path/to/comictagger"  # Path to Comictagger
$logFile = "/path/to/log/comictagger.log"  # Log file (create folder first)

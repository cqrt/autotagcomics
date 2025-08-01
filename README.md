Comic Auto-Tagger and Renamer Script for Windows

Monitors a directory, tags new comics with metadata, and renames them

Renaming template: {series} Vol.{volume} #{issue} ({year}) - Excludes volume if missing

CONFIGURATION

$watchFolder = "C:\path\to\comics"  # Directory to monitor

$comictagger = "C:\path\to\comictagger.exe"  # Path to Comictagger

$logFile = "C:\path\to\logs\Comictagger.log"  # Log file (create folder first)

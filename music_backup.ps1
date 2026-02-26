<#
    Script: music_backup.ps1
    Usage:
	powershell.exe -ExecutionPolicy Bypass -File "Path\To\music_backup.ps1" 

    Optional switches
	-DryRun    : simulates backup run without copying files
	-NoEmail   : suppresses email alerts
    -AltUser   : use alternate user credentials for network share access
    -Destination <path> : specify a custom destination root (UNC or drive); defaults to \\vault\Music

    Notes:
	- Requires folders.txt in the same directory
	- Logs are saved in the 'logs' subfolder
	- Email alerts use Gmail SMTP (App Password required if 2FA enabled)

    Version: 1.00
    Last Updated: 2026-02-26

    Changelog:
        0.0.1 - Initial PowerShell conversion from batch (https://www.ubackup.com/synchronization/robocopy-multiple-folders-6007-rc.html)
        0.1.0 - Added email alerts and logging
        0.1.3 - Added dry-run mode and interactive detection
        0.1.4 - Added command-line switches for -DryRun and -NoEmail
        0.1.5 - Improved logging format and summary
        0.2.0 - Added alternate user support for network share access
        0.2.1 - Enhanced email function - email credentials are now stored securely, the script will prompt for setup on first run
            - Added TLS 1.2 enforcement for SMTP connections
            - Improved folder list handling to ignore comments and empty lines
        0.2.2 - Added copied files tracking and attachment to email summary
        0.2.3 - Improved handling of spaces in paths and log file encoding (UTF-8 without BOM)
            - added more robust error handling and logging for directory creation and PSDrive mapping
        0.2.4 - Added per-run robocopy logs to extract copied files
            - improved filtering of copied files to exclude those matching excluded patterns
            - enhanced email summary with list of copied files and failed folders.
        1.0.0 - Added switch for custom destination path, updated documentation and renumbered changelog to reflect feature completeness.
            - This script is now considered stable for regular use.
            - Tagged as v1.00 in GitHub repository and published as Release
        1.0.1 - Added help display when run interactively with no parameters
        #>


param (
    [switch]$DryRun,
    [switch]$NoEmail,
    [switch]$AltUser,
    [string]$Destination = "\\vault\Music"   # base destination (UNC or drive letter)
)

# normalize double‑hyphen switches so that --DryRun etc behave like -DryRun
# this allows Unix-style invocation without breaking normal parameter binding
foreach ($raw in $args) {
    switch -Regex ($raw) {
        '^--DryRun$'    { $DryRun = $true }
        '^--NoEmail$'   { $NoEmail = $true }
        '^--AltUser$'   { $AltUser = $true }
        '^--Destination=(.+)' { $Destination = $Matches[1] }
    }
}

# Display help if run interactively with no parameters
if (($Host.Name -eq "ConsoleHost") -and -not $DryRun -and -not $NoEmail -and -not $AltUser -and $Destination -eq "\\vault\Music") {
    Write-Host "`n" @"
╔════════════════════════════════════════════════════════════════════╗
║                      Music Backup Script v1.00                     ║
╚════════════════════════════════════════════════════════════════════╝

DESCRIPTION:
  Synchronizes music folders from local/network sources to a 
  destination using Robocopy, with optional email alerts. Logs
  are saved in a subfolder, with timestamped filenames. 

USAGE:
  .\music_backup.ps1 [options]

OPTIONS:
  -DryRun               Simulate the backup without copying files
  -NoEmail              Suppress email notification
  -AltUser              Use alternate credentials for network share access
  -Destination <path>   Custom destination path (UNC or drive letter)
                        Default: \\vault\Music

EXAMPLES:
  .\music_backup.ps1                               # Run with defaults
  .\music_backup.ps1 -DryRun                       # Simulate run
  .\music_backup.ps1 -Destination '\\server\backup' # Custom destination
  .\music_backup.ps1 -AltUser -NoEmail -DryRun     # Alternate user, no mail, simulate

REQUIREMENTS:
  • folders.txt in the script directory (list of source folders)
  • Gmail SMTP requires App Password if 2FA is enabled

FIRST RUN:
  Run this script from a powershell prompt, you'll be prompted to create an encrypted
  credential file for email authentication.

For more details, see the comment block at the top of this script.
"@
    exit 0
}

# Check if running interactively and credential file doesn't exist

$emailCredPath = "$PSScriptRoot\email_cred.xml"
if (($Host.Name -eq "ConsoleHost") -and !(Test-Path $emailCredPath)) {
    Write-Host "`n=== First Run Email Credential Setup ==="
    Write-Host "No email credential file found. Let's create one."
    Write-Host "Note: Use an App Password if 2FA is enabled on your Gmail account"
    Write-Host "Generate one at: https://myaccount.google.com/apppasswords"
    
    $emailUser = Read-Host "Enter Gmail address"
    $emailPass = Read-Host "Enter Gmail App Password" -AsSecureString
    
    # Create credential object and export
    $emailCred = New-Object PSCredential($emailUser, $emailPass)
    $emailCred | Export-Clixml -Path $emailCredPath
    
    # Set strict file permissions
    $acl = Get-Acl $emailCredPath
    $acl.SetAccessRuleProtection($true, $false)
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        "$env:USERDOMAIN\$env:USERNAME", "Read", "Allow")
    $acl.AddAccessRule($rule)
    Set-Acl $emailCredPath $acl
    
    Write-Host "Credential file created and secured at: $emailCredPath`n"
}

# If alternate user is requested, map network share
# Need to run the following from an interactive PowerShell console first:
# 
# $securePassword = Read-Host "Enter password" -AsSecureString
# $securePassword | Export-Clixml -Path "D:\MusicBackup\music_backup_password.xml"

if ($AltUser) {
    $credPath = "D:\MusicBackup\music_backup_password.xml"
    $sharePath = "\\vault\Music"

    function Get-FreeDriveLetter {
        param([string]$Preferred = "Z")
        $used = @()
        Get-PSDrive -PSProvider FileSystem | ForEach-Object { $used += $_.Name.ToUpper() }
        [System.IO.DriveInfo]::GetDrives() | ForEach-Object { $used += $_.Name.Substring(0,1).ToUpper() }
        $candidates = @()
        if ($Preferred) { $candidates += $Preferred.Substring(0,1).ToUpper() }
        for ($i=[int][char]'Z'; $i -ge [int][char]'D'; $i--) { $candidates += [char]$i }
        foreach ($letter in $candidates) { if (-not ($used -contains $letter)) { return $letter } }
        return $null
    }

    $mappedDrive = Get-FreeDriveLetter -Preferred "Z"
    if (-not $mappedDrive) {
        Write-Error "No free drive letters available. Aborting."
        exit 1
    }

    if (-not (Test-Path $credPath)) {
        Write-Error "Credential file not found at $credPath. Run the setup command to create it."
        exit 1
    }
    $cred = Import-Clixml -Path $credPath

    New-PSDrive -Name $mappedDrive -PSProvider FileSystem -Root $sharePath -Credential $cred -Persist | Out-Null
    $destRoot = "${mappedDrive}:\"
}

# Helper function for UTF-8 logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Path
    )
    # UTF-8 encoding without BOM
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    
    # Ensure directory exists
    $logDir = Split-Path -Parent $Path
    if (!(Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Write to file using .NET directly for consistent UTF-8 handling
    if (Test-Path $Path) {
        [System.IO.File]::AppendAllText($Path, $Message + [Environment]::NewLine, $utf8NoBom)
    } else {
        [System.IO.File]::WriteAllText($Path, $Message + [Environment]::NewLine, $utf8NoBom)
    }
}

# Set script root and log directory
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logDir = Join-Path $scriptRoot "logs"
if (!(Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory }

# Generate timestamped log file
$timestamp = Get-Date -Format "dd-MM-yyyy_HH-mm-ss"
$logFile = "$logDir\robocopy_$timestamp.txt"

# Define destination root (may be overridden by -Destination parameter)
$destRoot = $Destination

# Define exclusions
$excludeFiles = @("*.bak", "*.backup", "syncthing*.*")      # File patterns to exclude
$excludeDirs = @(".stfolder")                               # Directory names to exclude

# Detect if running interactively
$isInteractive = ($Host.Name -eq "ConsoleHost")

# Show configuration summary
if ($isInteractive) {
    Write-Host "`n=== Script Configuration ==="

    if ($DryRun -and $NoEmail) {
        Write-Host "Dry-run mode enabled: Robocopy will simulate actions only"
        Write-Host "Email alerts are suppressed"
    } elseif ($DryRun) {
        Write-Host "Dry-run mode enabled: Robocopy will simulate actions only"
    } elseif ($NoEmail) {
        Write-Host "Email alerts are suppressed"
    } else {
        Write-Host "Live mode: Robocopy will perform actual sync and send email alerts"
    }

    Write-Host "`nSource list file : $scriptRoot\folders.txt"
    Write-Host "Destination root  : $destRoot"
    Write-Host "Log file path     : $logFile"
    Write-Host "=============================`n"
}

# Initialize counters
$copiedCount = 0
$skippedCount = 0
$failedCount = 0
$failedFolders = @()

# NEW: track copied files across runs
$copiedFiles = @()
$runIndex = 0

# Start log
Write-Log "=== Robocopy Sync Started: $timestamp ===" $logFile

# Read folders.txt and ignore comment lines
# Each line is treated as a complete path (spaces in paths are fully preserved)
$rawFolders = @(Get-Content "$scriptRoot\folders.txt" -Raw -ErrorAction SilentlyContinue)
$folderList = @()
if ($rawFolders) {
    $rawFolders -split "`n" | ForEach-Object {
        $line = $_.Trim()
        # Skip empty lines and comments
        if ($line -and -not $line.StartsWith("#")) {
            $folderList += $line
        }
    }
}

if ($folderList.Count -eq 0) {
    Write-Error "No source folders found in $scriptRoot\folders.txt"
    exit 1
}

Write-Log "Found $($folderList.Count) source folders to process" $logFile

foreach ($sourcePath in $folderList) {
    # Log the exact path being processed (for debugging spaces)
    Write-Log "DEBUG: Reading path='$sourcePath' (Length: $($sourcePath.Length) chars)" $logFile
    
    if (Test-Path -LiteralPath $sourcePath) {
        $runIndex++
        # Extract relative path (strip drive letter and colon, e.g., "D:" -> "")
        $relativePath = if ($sourcePath -match '^[A-Za-z]:\\(.*)$') { $matches[1] } else { $sourcePath.TrimStart('\') }
        $destPath = Join-Path $destRoot $relativePath

        # Ensure destination directory exists before running robocopy
        # (robocopy may fail with error 16 if it can't create the destination structure)
        if (!(Test-Path -LiteralPath $destPath)) {
            try {
                New-Item -ItemType Directory -Path $destPath -Force | Out-Null
                Write-Log "Created destination directory: $destPath" $logFile
            } catch {
                Write-Log "ERROR: Failed to create destination directory: $destPath - $_" $logFile
                $failedCount++
                $failedFolders += $sourcePath
                continue
            }
        }

        # Per-run robocopy log so we can extract copied files
        $runLog = Join-Path $logDir ("robocopy_run_{0}_{1}.txt" -f $timestamp, $runIndex)
        
        # Log the paths being used for debugging
        Write-Log "Processing: Source=$sourcePath | Dest=$destPath" $logFile

        # Build robocopy arguments as clean array (Start-Process handles quoting)
        $robocopyArgs = @(
            $sourcePath,        # Do NOT manually quote - let Start-Process handle it
            $destPath,          # Same here
            "/MT:16",
            "/E",
            "/V",
            "/FP",
            "/TS",
            "/LOG+:$runLog"     # Log path may have spaces, but /LOG+ treats the whole argument
        )

        # Add exclude filters - use /XF for files, /XD for directories
        foreach ($ex in $excludeFiles) {
            $robocopyArgs += "/XF"
            $robocopyArgs += $ex
        }
        foreach ($ex in $excludeDirs) {
            $robocopyArgs += "/XD"
            $robocopyArgs += $ex
        }

        if ($DryRun) {
            $robocopyArgs += "/L"
        }

        # Use call operator (&) instead of Start-Process for better handling of spaces in paths
        Write-Log "Executing: robocopy.exe $($robocopyArgs -join ' ')" $logFile
        & robocopy.exe @robocopyArgs
        $exitCode = $LASTEXITCODE

        # Robocopy exit codes:
        # 0 = No files copied (but success - no changes needed)
        # 1 = Files copied successfully
        # 2 = Extra files or directories detected (also success indicator)
        # 3 = Files copied + extra files/dirs
        # 4 = Mismatched files/dirs detected
        # 8+ = Serious errors - some files not copied
        
        # Extract only files that were actually copied (not already existing)
        if (Test-Path $runLog) {
            # Robocopy with /V outputs action indicators: "New File", "Newer", "named", etc.
            Get-Content $runLog | ForEach-Object {
                $line = $_
                # Only match lines that indicate a file/dir was copied: "New File", "Newer", or "named" action
                if ($line -match '(New File|Newer|named)' -and $line -match '[A-Za-z]:\\') {
                    # Extract just the path (last tab-separated field)
                    $parts = $line -split '\t'
                    $copyPath = $parts[-1].Trim()
                    if ($copyPath -and $copyPath -notmatch '^\d+$') {  # Avoid capturing lone numbers/whitespace
                        # Filter out any paths in excluded directories
                        $shouldExclude = $false
                        foreach ($exDir in $excludeDirs) {
                            # match the directory name plus trailing backslash
                            if ($copyPath -match [regex]::Escape("$exDir\")) {
                                $shouldExclude = $true
                                break
                            }
                        }
                        # Filter out files matching excluded patterns
                        if (-not $shouldExclude) {
                            foreach ($exFile in $excludeFiles) {
                                $filePattern = $exFile -replace '\*', '.*' -replace '\.', '\.'
                                if ($copyPath -match $filePattern) {
                                    $shouldExclude = $true
                                    break
                                }
                            }
                        }
                        if (-not $shouldExclude) {
                            $copiedFiles += $copyPath
                        }
                    }
                }
            }
        }

        # Log result based on robocopy exit code and files found
        # Exit codes 1-3 indicate files/directories were copied
        # Exit code 0 means no files needed copying (already in sync)
        if ($exitCode -le 4) {
            if ($exitCode -in @(1, 2, 3)) {
                # Exit codes 1, 2, or 3 mean files or directories were copied
                $copiedCount++
                Write-Log "Copied: $sourcePath -> $destPath (Exit Code: $exitCode, files/folders copied)" $logFile
            } elseif ($exitCode -eq 0 -or $exitCode -eq 4) {
                # Exit code 0 = no files copied / already in sync
                # Exit code 4 = mismatched files (but no actual failures)
                Write-Log "Synced: $sourcePath -> $destPath (Exit Code: $exitCode, no new files)" $logFile
            }
        } else {
            $failedCount++
            $failedFolders += $sourcePath
            Write-Log "FAILED: $sourcePath -> $destPath (Exit Code: $exitCode)" $logFile
        }
    } else {
        $skippedCount++
        Write-Log "Skipped (not found): $sourcePath" $logFile
    }
}

Write-Log "=== Robocopy Sync Completed: $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss') ===" $logFile
Write-Log "Summary: Copied=$copiedCount, Skipped=$skippedCount, Failed=$failedCount" $logFile

# Show summary in console
if ($isInteractive) {
    Write-Host "`nRobocopy Sync Summary:"
    Write-Host "Copied folders: $copiedCount"
    Write-Host "Skipped (not found): $skippedCount"
    Write-Host "Failed copies: $failedCount"
    Write-Host "Log saved to: $logFile`n"
}

# Prepare email summary
$body = @"
The scheduled Robocopy sync completed on $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss').

Summary:
- Copied folders: $copiedCount
- Skipped (not found): $skippedCount
- Failed copies: $failedCount
"@

if ($failedCount -gt 0) {
    $body += "`nFailed folders:`n" + ($failedFolders -join "`n") + "`n"
}

$body += "`nSee attached log for full details."

# NEW: write combined copied-files attachment
$copiedFilesFile = Join-Path $logDir ("copied_files_{0}.txt" -f $timestamp)
if ($copiedFiles.Count -gt 0) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($copiedFilesFile, $copiedFiles, $utf8NoBom)
} else {
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($copiedFilesFile, "No files copied during this run." + [Environment]::NewLine, $utf8NoBom)
}

# Send email alert via Gmail unless -NoEmail set
if (-not ($NoEmail)) {
    $smtpServer = "smtp.gmail.com"
    $smtpPort = 587
    
    # Load credentials from file
    if (Test-Path $emailCredPath) {
        $emailCred = Import-Clixml -Path $emailCredPath
        $from = $emailCred.UserName
        $to = "rbyrnes@gmail.com"  # Keep your existing recipient
        
        $subjectPrefix = if ($DryRun) { "[DRY RUN] " } else { "" }
        if ($failedCount -gt 0) {
            $subject = "$subjectPrefix[FAILED] Robocopy Sync"
        } else {
            $subject = "$subjectPrefix[SUCCESS] Robocopy Sync"
        }
        
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

        try {
            # Attach both the main log and the copied-files list
            $attachments = @($logFile, $copiedFilesFile)
            Send-MailMessage -From $from -To $to -Subject $subject -Body $body `
                -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $emailCred `
                -Attachments $attachments
        }
        catch {
            Write-Error "Failed to send email: $_"
        }
    }
    else {
        Write-Error "Email credential file not found at: $emailCredPath"
    }
}

# Cleanup mapped drive
if ($AltUser -and $mappedDrive) {
    try {
        Remove-PSDrive -Name $mappedDrive -Force -ErrorAction Stop
    } catch {
        Write-Log "Warning: failed to remove PSDrive ${mappedDrive}: $_" $logFile
    }
}

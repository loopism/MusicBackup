<#
    Script: music_backup.ps1
    Usage:
	powershell.exe -ExecutionPolicy Bypass -File "Path\To\music_backup.ps1" 

    Optional switches
	-DryRun    : simulates backup run without copying files
	-NoEmail   : suppresses email alerts
    -AltUser   : use alternate user credentials for network share access

    Notes:
	- Requires folders.txt in the same directory
	- Logs are saved in the 'logs' subfolder
	- Email alerts use Gmail SMTP (App Password recommended)

    Version: 1.5
    Last Updated: 2025-11-05

    Changelog:
        1.0 - Initial PowerShell conversion from batch (https://www.ubackup.com/synchronization/robocopy-multiple-folders-6007-rc.html)
        1.1 - Added email alerts and logging
        1.2 - Added dry-run mode and interactive detection
        1.3 - Added command-line switches for -DryRun and -NoEmail
        1.4 - Improved logging format and summary
        1.5 - Added alternate user support for network share access
#>


param (
    [switch]$DryRun,
    [switch]$NoEmail,
    [switch]$AltUser
)

# Default destination root
$destRoot = "\\vault\Music"

# If alternate user is requested, map network share
# Need to run the following from an interactive PowerShell console first:
# 
# $securePassword = Read-Host "Enter password" -AsSecureString
# $securePassword | Export-Clixml -Path "D:\MusicBackup\music_backup_password.xml"

if ($AltUser) {
    $credPath = "D:\MusicBackup\music_backup_password.xml"
    $username = "OtherUser"  # or  "DOMAIN\OtherUser" if not local
    $sharePath = "\\vault\Music"
    $mappedDrive = "MusicShare"

    if (Test-Path $credPath) {
        $securePassword = Import-Clixml -Path $credPath
        $cred = New-Object System.Management.Automation.PSCredential ($username, $securePassword)

        # Map the share using stored credentials
        New-PSDrive -Name $mappedDrive -PSProvider FileSystem -Root $sharePath -Credential $cred -Persist | Out-Null
        $destRoot = "$mappedDrive:\"
    } else {
        Write-Error "Credential file not found at $credPath. Run the setup command to create it."
        exit 1
    }
}


# Helper function for UTF-8 logging
function Write-Log {
    param (
        [string]$Message,
        [string]$Path
    )
    $Message | Out-File -FilePath $Path -Append -Encoding utf8
}

# Set script root and log directory
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$logDir = Join-Path $scriptRoot "logs"
if (!(Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory }

# Generate timestamped log file
$timestamp = Get-Date -Format "dd-MM-yyyy_HH-mm-ss"
$logFile = Join-Path $logDir "robocopy_$timestamp.txt"

# Define destination root
$destRoot = "\\vault\Music"

# Define exclusions
$excludes = @("*.bak", "*.backup", ".stfolder", "syncthing*.*")

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

# Start log
Write-Log "=== Robocopy Sync Started: $timestamp ===" $logFile

# Read folders.txt
$folderList = Get-Content "$scriptRoot\folders.txt"

foreach ($sourcePath in $folderList) {
    if (Test-Path $sourcePath) {
        $relativePath = $sourcePath.Substring(3)
        $destPath = Join-Path $destRoot $relativePath

        $excludeArgs = $excludes | ForEach-Object { "/XF `"$($_)`"" }
        $robocopyArgs = @(
            "`"$sourcePath`"", "`"$destPath`"",
            "/MT", "/E", "/XO", "/FFT", "/V", "/NDL", "/NFL"
        )

        if ($DryRun) {
            $robocopyArgs += "/L"
        }

        $robocopyArgs += $excludeArgs
        $cmd = "robocopy " + ($robocopyArgs -join " ")
        $result = Invoke-Expression $cmd
        $exitCode = $LASTEXITCODE

        if ($exitCode -lt 8) {
            $copiedCount++
            Write-Log "Copied: $sourcePath → $destPath (Exit Code: $exitCode)" $logFile
        } else {
            $failedCount++
            $failedFolders += $sourcePath
            Write-Log "FAILED: $sourcePath → $destPath (Exit Code: $exitCode)" $logFile
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
# Send email alert via Gmail unless -NoEmail set
if (-not ($NoEmail)) {
    $smtpServer = "smtp.gmail.com"
    $smtpPort = 587
    $from = "serverlogging.allservers@gmail.com"
    $to = "rbyrnes@gmail.com"
    $subjectPrefix = if ($dryRun) { "[DRY RUN] " } else { "" }
    if ($failedCount -gt 0) {
        $subject = "$subjectPrefix❌ Robocopy Sync FAILED"
    } else {
        $subject = "$subjectPrefix✅ Robocopy Sync Completed"
    }
}

# Secure password (use App Password if 2FA is enabled)
$securePassword = ConvertTo-SecureString "igad jcdx amat axyf" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($from, $securePassword)

Send-MailMessage -From $from -To $to -Subject $subject -Body $body `
    -SmtpServer $smtpServer -Port $smtpPort -UseSsl -Credential $cred `
    -Attachments $logFile

# Cleanup mapped drive
if ($AltUser) {
    Remove-PSDrive -Name "MusicShare"
}

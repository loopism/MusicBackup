A Powershell script that takes a list of folders and copies them to a remote network share.
This script can either be run interactively or from Task Scheduler.
It uses robocopy to copy files that are not on the remote share or have been changed locally
and will log the output in a file then email a summary.

Email credentials are stored securely using Windows Data Protection API (DPAPI), the script will prompt for
credentials on the first run.

Interactive usage:
	powershell.exe -ExecutionPolicy Bypass -File "Path\To\music_backup.ps1" 

Optional switches
-DryRun    : simulates backup run without copying files
-NoEmail   : suppresses email alerts
-AltUser   : use alternate user credentials for network share access

Notes:
- Requires folders.txt in the same directory to define which folders are copied to the destination
- Logs are saved in the 'logs' subfolder
- Email alerts use Gmail SMTP (App Password recommended)

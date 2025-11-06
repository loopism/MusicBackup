A Powershell script that takes a list of folders and copies them to a remote network share.
This script can either be run interactively or from Task Scheduler.
It uses robocopy to copy files that are not on the remote share or have been changed locally
and will log the output in a file then email a summary.

Email credentials are stored securely using Windows Data Protection API (DPAPI), the script will prompt for
credentials on the first run. 
There are some caveats:
  - The credential file exists (created during the first interactive run)
  - The script is running under the same Windows user account that created the credential file
  - The script is running on the same computer where the credential file was created

DPAPI uses the user's login credentials as part of the encryption key, so the credential file can
only be decrypted by the same user on the same machine.

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

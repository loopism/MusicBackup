A Powershell script that takes a list of folders and copies them to somewhere else - usually a remote network share.
This script can either be run interactively or from Task Scheduler.
It uses robocopy to copy only files that are not on the remote share or have been changed locally
and will log the output in a file then email a summary.

Running the script without any switches in an interactive session will display a built‑in help screen summarising usage and examples.

Email credentials are stored securely using Windows Data Protection API (DPAPI), DPAPI uses the user's login credentials 
as part of the encryption key.

There are some caveats:<br>
  - The credential file exists (if not, it is created during the first interactive run)
  - The script must under the same Windows user account that created the credential file
  - The script must run on the same computer where the credential file was created

Interactive usage:

	powershell.exe -ExecutionPolicy Bypass -File "Path\To\music_backup.ps1" -OptionalSwitches

Optional switches<br>
-DryRun      : simulates backup run without copying files<br>
-NoEmail     : suppresses email alerts<br>
-AltUser     : use alternate user credentials for network share access<br>
-Destination : specify a custom destination path (UNC or drive); defaults to \\vault\Music<br>

Notes:
- Requires folders.txt in the same directory as the script to define which folders are copied to the destination
- Logs are saved in the 'logs' subfolder (per‑run logs also created for file extraction)
- Email alerts use Gmail SMTP (App Password required if the associated Google account uses 2FA)

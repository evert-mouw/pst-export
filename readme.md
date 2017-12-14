# export-pst

## About

Export a specific mailbox on an Exchange server to a PST file. You can use it to create backups, to copy data or whatever.

Tested with Exchange 2016.

## Installing

You might want to copy these files to e.g. `C:\pst-export`

Edit all the files and replace user names, server names, path names, codes and email addresses specific for your environment.

Afer some testing you might want to enable a scheduled task. Example XML code is included.

For mail notifications you might want to copy `bmail.exe` in the same folder. Bmail is a command-line mailer; see [bmail](http://retired.beyondlogic.org/solutions/cmdlinemail/cmdlinemail.htm).

## Importing

To *import* a PST file, you need another set of commands.

You are Administrator. Open a PowerShell with the Exchange extensions loaded. First, make sure to assign the role. Then you enable a mailbox, and finally you can create the import request.

	New-ManagementRoleAssignment -Role "Mailbox Import Export" -User Administrator
	Enable-Mailbox -Identity SOMEUSER
	New-MailboxImportRequest -Mailbox SOMEUSER -FilePath "\\localhost\C$\BACKUPS\SOMEUSER.pst"

To show the progress of the work, do: `Get-MailboxImportRequest`

## Author

	Evert Mouw <post@evert.net>

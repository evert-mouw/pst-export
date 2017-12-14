@echo off
rem Creates a PST export
rem by Evert Mouw

C:
cd c:\pst-export

rem Move the last successfull export to the last_success folder.
if exist USER1.pst (
	del last_success\*.* /q
	move *.pst last_success\
	move *.ics last_success\
	move date.txt last_success\
) else (
	rem Send me an email about this!
	bmail -h -s SMTP.EXAMPLE.COM -t ADMINEMAIL -f ADMINEMAIL -a "#! PST export failed" -b "Please check C:\pst-export"
)

echo "Date of backup start:" > date.txt
date /t >> date.txt
time /t >> date.txt

rem First get the ICS files (calendar publishing is enabled by these users).
rem user = USER1
curl --silent --insecure -O https://EXCHANGESERVER/owa/calendar/USERSPECIFICCODE/calendar.ics
ren reachcalendar.ics USER1.ics
rem user = USER2
curl --silent --insecure -O https://EXCHANGESERVER/owa/calendar/USERSPECIFICCODE/calendar.ics
ren reachcalendar.ics USER2.ics

rem Now do the PST export -- a Powershell script is needed though.
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; .\pst-export.ps1"

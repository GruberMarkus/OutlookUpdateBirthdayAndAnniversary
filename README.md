WHAT IS IT?

¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Have you ever forgot the birthday of a friend, colleague or customer because Outlook did not remind you about it?
The reason for this may be that you set or updated birthdays and anniversariers on your mobile or in webmail and not in the full Outlook client on Windows. Usually, the reminders then do not appear in your Outlook calendar and your Outlook birthday calendar.

This Visual Basic Script (VBS) goes through your contacts on Outlook on Windows and re-writes all birthday and anniversary dates, making them appear in your calendars. Never miss a birthday or anniversary reminder again!

HOW DO I RUN THIS SCRIPT?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Option 1
	Double click "OutlookUpdateContactBirthdayAndAnniversary.vbs"
Option 2
	In a command line, navigate to the folder the file is stored in and execute "cscript.exe OutlookUpdateContactBirthdayAndAnniversary.vbs"


HOW OFTEN SHOULD I RUN THIS SCRIPT?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
You can run this script as often as you want. I recommend running it regularly to ensure you never miss a birthday or anniversary reminder again. If you are an advanced user, you can automate the script execution.


NOTES
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
.) This script is very basic, more a demonstrator than a finished product. Especially, it misses error handling routines.
.) This script works with Outlook 2013 and up. Only the Windows platform is supported.
.) The script works for contacts stored in any folder marked to contain contact items. The folders do not need to be marked as Address Books.
.) Hidden but writable folders are excluded from updates per default. Read-only folders are excluded from updates, of course.
.) The script does not clean up birthday or anniversary calendar entries that were created manually or belong to contacts that no longer exist.
.) The script supports multiple accounts and data files within one profile. Contacts are updated for all accounts and data files within a profile.
.) If Outlook is not started, the script starts it automatically (in the background).
.) The script supports multiple Outlook profiles:
..) If Outlook is already running, the currently active profile is used.
..) If Outlook is started by the script, it will per default ask for the profile to use (no matter how you have configured this setting in Outlook). Advanced users may change this option in the script.


WHAT ABOUT THE LICENSE?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
There is no license attached to this little helper program. Basically, you can do what you want with it.
There is no commercial support available. I support the tool in my spare time, so there is no guarantee here, too.


HOW CAN I MODIFY IT OR COMPILE IT MYSELF?
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
The tool is just a few lines of Visual Basic Script (VBS) code, available in the script file as clear text.
For people with VBS knowledge, there are some configurable options at the top of the source code.


THE PROGRAM DOES NOT WORK
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
In this case, please contact me to schedule a remote session with access to your client to analyze the problem.

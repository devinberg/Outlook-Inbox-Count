Outlook Inbox Count
===================
A Python script to record the number of items in your Outlook inbox to a text file.

Setup
======
For this script to work, you need to do the following.

1. Install [Pywin32](http://sourceforge.net/projects/pywin32/) for Python 3.x
2. Enable the script to load when you start Outlook using VBA scripting. In Outlook press ALT+F11 to open the Visual Basic editor then paste in the following code.
        Private Sub Application_Startup()
		    'Edit path as appropriate.
            Shell ("pythonw C:\outlook.py")
        End Sub
3. Kill the script when you exit Outlook. Paste the following in the editor.
        Private Sub Application_Quit()
            Shell ("taskkill /F /IM pythonw.exe")
        End Sub
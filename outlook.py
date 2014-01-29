import win32com.client
import time
import datetime

connected = 1
# Start the loop
while connected == 1 :
    # Connect to Outlook, which has to be running
    try:
        outlook = \
    win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except:
        print ('Could not open Outlook')
        connected = 0
    connected = 1

    # Open the text file.
    f = open('C:/inbox.txt', 'a')

    # Count the number of messages in the inbox
    inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)
    messages = inbox.Items

    # Get the time.
    ts = time.time()
    st = datetime.datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')


    # Write the message number and time to text file.
    f.write(st + ',' + str(messages.Count) +'\n')
    f.close()
    time.sleep(1800)

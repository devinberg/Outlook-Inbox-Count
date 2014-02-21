import win32com.client
import time
from numpy import *
from pylab import *
import matplotlib.dates as mdates
import csv
from datetime import datetime

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
    f = open('C:/inbox.txt', 'a') # Change to desired directory.

    # Count the number of messages in the inbox
    inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)
    messages = inbox.Items

    # Get the time.
    ts = time.time()
    st = datetime.fromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')


    # Write the message number and time to text file.
    f.write(st + ',' + str(messages.Count) +'\n')
    f.close()
	
	# Read text file and create plot.
    data = genfromtxt('C:/inbox.txt', delimiter=',', dtype = str)

    dates = [row[0] for row in data]
    count = [row[1] for row in data]

    dates[:] = [datetime.strptime(x, "%Y-%m-%d %H:%M:%S") for x in dates]

    gca().xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=(1),
                                                    interval=1))
    gca().xaxis.set_minor_formatter(mdates.DateFormatter('%d\n%a'))
    gca().xaxis.set_major_formatter(mdates.DateFormatter('\n\n\n%b\n%Y'))
    gca().xaxis.set_major_locator(mdates.MonthLocator())
    plot(dates,count)
    gcf().autofmt_xdate()
    xlabel('Date')
    ylabel('Inbox Count')
    tight_layout()
    savefig('C:\\Users\\path\\to\\folder\\inbox.png') # Change to desired directory.
	matplotlib.pyplot.close("all") # Close the figure so it doesn't overwrite every time.
    time.sleep(1800) # Wait 30 mins before checking the inbox again.

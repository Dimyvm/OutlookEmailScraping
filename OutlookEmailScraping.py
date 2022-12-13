import win32com.client
from datetime import date, timedelta, datetime
import win32ui
from localStorage import *
from articleClass import *


# def read whene it was the last time this script was executed
lastRun = lastDateRun()
print(f'This script was run successfully for the last time on {lastRun}')

# def read outlook emails
outlook = win32com.client.dynamic.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
# "6" refers to the index of a folder - in this case,
# "5" = Verzonden items
messages = inbox.Items  # all mails from inbox


# def filter emails
messages.Sort("[ReceivedTime]", True)  # sort messages on date
lastRunPlusOne = datetime.strptime(lastRun, "%Y-%m-%d") + timedelta(days=1)
messagesToday = messages.Restrict(
    "[ReceivedTime] >= '" + lastRunPlusOne.strftime('%d/%m/%Y %H:%M %p')+"'")
print(
    f"There are {messagesToday.count} messages collected from {lastRun} to today")

# read the data from each filtered email
aantal = 0
for message in messagesToday:

    if message.subject.startswith('SIEMENS - Update'):
        aantal += 1
        body_title = message.subject
        body_content = message.body
        # sendDate = message.ReceivedTime
        sendDate = message.SentOn.strftime("%d-%m-%y")

        # if aantal == 1:
        print(f"Send date : {sendDate}")
        print(body_title)
        # print(body_content)
        lines = body_content.splitlines()
        print(lines[9])  # bestelbon

        for line in lines:
            if "Klantartikel" not in line:
                if line.startswith("ArtikelID"):
                    articleId = line.split()
                    print(articleId[1])
                if line.startswith("Bevestigde leverdatum"):
                    deliveryDate = line.split()
                    print(deliveryDate[2])
                if line.startswith("Bevestigd aantal"):
                    number = line.split()
                    print(number[2])


print(f"There are {aantal} messages found with the right subject")


# def open excel

# delete each record where the deliverydate == to the date of Today

# Write each article and deliverydate in a excel record

# save and close excel


# if script is executed successfully then show windows messageBox
if aantal > 0:

    win32ui.MessageBox("""Successfully executed!""", "Siemens email scraping")

    # write the date whene this script was executed
    updateLastDateRun()

else:
    # Else --script has a error then send a email
    print("script has a error then send a email")

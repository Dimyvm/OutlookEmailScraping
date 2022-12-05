import win32com.client
from datetime import date, timedelta
import win32ui
from localStorage import *



# FLOW
# def read whene it was the last time this script was executed
lastDateRun()
# if script is executed successfully then show windows messageBox
# def read outlook emails
# def filter emails
# read the data from each filtered email
# def open excel
# delete each record where the deliverydate == to the date of Today
# Write each article and deliverydate in a excel record
# save and close excel
# write the date whene this script was executed
# IF script has a error then send a email


outlook = win32com.client.dynamic.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# "6" refers to the index of a folder - in this case,
# "5" = Verzonden items
inbox = outlook.GetDefaultFolder(6)


messages = inbox.Items  # all mails from inbox

messages.Sort("[ReceivedTime]", True)  # sort messages on date


yesterday = date.today() - timedelta(days=1)  # get dat yesteday
messagesToday = messages.Restrict("[ReceivedTime] >= '" + yesterday.strftime(
    '%d/%m/%Y %H:%M %p')+"'")  # filter only the messages of today

print(f"There are {messagesToday.count} messages today")
# message = messages.GetLast()
aantal = 0
for message in messagesToday:

    if message.subject.startswith('SIEMENS - Update'):
        aantal += 1
        body_title = message.subject
        body_content = message.body
        # sendDate = message.ReceivedTime
        sendDate = message.SentOn.strftime("%d-%m-%y")

        if aantal == 1:
            print(f"Send date : {sendDate}")
            print(body_title)
            # print(body_content)
            lines = body_content.splitlines()
            print(lines[9])  # bestelbonn
            for line in lines:
                if "Klantartikelnummer " not in line:
                    if line.startswith("Artikelnummer") or line.startswith("Bevestigde leverdatum"):
                        print(line)

print(f"There are {aantal} messages found with the right subject")
win32ui.MessageBox("""Successfully executed!""", "Siemens email scraping")

updateLastDateRun()



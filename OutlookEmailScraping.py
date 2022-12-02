import win32com.client
from datetime import date, timedelta

outlook = win32com.client.dynamic.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# "6" refers to the index of a folder - in this case,
# 5 = Verzonden items
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
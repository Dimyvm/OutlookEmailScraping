import win32com.client

outlook = win32com.client.dynamic.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# "6" refers to the index of a folder - in this case,
# 5 = Verzonden items
inbox = outlook.GetDefaultFolder(6)

# all mails from inbox
messages = inbox.Items

messages.Sort("[ReceivedTime]", True)
# last30MinuteMessages = messages.Restrict("[ReceivedTime] >= '" +last30MinuteDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")
print(f"There are {messages.count} messages")
# message = messages.GetLast()
aantal = 0
for message in messages:

    if message.subject.startswith('SIEMENS - Update'):
        aantal += 1
        body_title = message.subject
        body_content = message.body
        sendDate = message.SentOn.strftime("%d-%m-%y")
        # sendDate = message.ReceivedTime

        if aantal == 1:
            print(f"Send date : {sendDate}")
            print(body_title)
            # print(body_content)
            lines = body_content.splitlines()
            print(lines[9])  # bestelbonn
            for line in lines:
                # f line.startswith("ArtikelID") or line.startswith("Bevestigde leverdatum"):
                if line.startswith("Artikelnummer") or line.startswith("Bevestigde leverdatum"):
                    print(line)
print(f"There are {aantal} messages found with the right subject")

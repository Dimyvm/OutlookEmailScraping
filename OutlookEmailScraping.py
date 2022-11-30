import win32com.client

outlook = win32com.client.dynamic.Dispatch(
    "Outlook.Application").GetNamespace("MAPI")

# "6" refers to the index of a folder - in this case,
# 5 = Verzonden items
inbox = outlook.GetDefaultFolder(6)

# the inbox. You can change that number to reference
# any other folder
messages = inbox.Items
print(f"There are {messages.count} messages")
# message = messages.GetLast()
aantal = 0
for message in messages:

    if message.subject.startswith('SIEMENS - Update van uw bestelling'):
        aantal += 1
        body_title = message.subject
        body_content = message.body
        if aantal == 1:
            print(body_title)
            # print(body_content)
            lines = body_content.splitlines()
            print(lines[9])  # bestelbonn
            for line in lines:
                if line.startswith("ArtikelID") or line.startswith("Bevestigde leverdatum"):
                    print(line)
print(f"There are {aantal} messages found with the right subject")
# body_content = message.body

# print(body_content)

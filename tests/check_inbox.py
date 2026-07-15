import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

for i, message in enumerate(messages):
    if i >= 5:
        break
    print(message.ReceivedTime, "-", message.Subject)
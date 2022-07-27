from getpass import getpass, getuser
import imaplib
from email import message_from_bytes
from pathlib import Path
from tkinter.filedialog import askdirectory
from os import chdir

'''Mail Addresses downloader - Nicola Moro - 2022
Simply download addresses from your mail service and store them'''


def formatAddress(string):

    newString = string

    try:
        if "<" in newString:
            newString = newString[newString.index("<")+1:newString.index(">")]
            return newString

        for char in string:
            if char != " ":
                newString = string[string.index(char):]
                break
        return newString

    except:
        return newString


print(
    '''\n| Mail addresses downloader - Moro Nicola - 2022
| Please insert the requested informations to retrieve all the contacts in your mail boxes\n'''
)

hostUser = input("> Type in the IMAP server: ")
portUser = input("> Type in the IMAP port number: ")

connectionClient = None


print("\n| Connecting...")
try:
    connectionClient = imaplib.IMAP4_SSL(host=hostUser, port=int(portUser),timeout=10)
    print(
        f"| Connection with {connectionClient.host} on port {connectionClient.port} established.")
except Exception as e:
    print(
        f"| Connection refused: ", e.args)
    print("| Please check you params and try again")
    quit()


print("\n| Please login to your account")
userMail = input("> Username: ")
userPassword = getpass(prompt="> Password: ")

print("| Trying to login...")
try:
    connectionClient.login(userMail, userPassword)
    print("| Login successful")
except Exception as e:
    print("| Login failed:", e.args)
    print("| Please check your credentials and try again. A specific app password may be required for the script to work.")
    quit()

while True:

    print("\n| These are the available folders in your account:")
    resp_code, directories = connectionClient.list()

    for n in directories:
        print(f"| ~ {n.decode()}")

    choice = input("\n> Please select the folder to download the addresses: ")

    typ, messages = connectionClient.select(f'"{choice}"')
    resp, mails = connectionClient.search(None, "ALL")
    mailsId = mails[0].decode().split()
    addresses = []

    print("| Job started. This could take a while, depending on how many message there are in the mailbox...\n")
    for n in mailsId:
        resp, maildata = connectionClient.fetch(n, "(RFC822)")
        message = message_from_bytes(maildata[0][1])
        sender = message.get("From")
        receiver = message.get("To")
        if sender not in addresses:
            addresses.append(sender)
        if receiver not in addresses:
            addresses.append(receiver)
        print(f"From {sender} to {receiver}")

    print("\n| Reading end. Trying to normalize addresses...")
    for i, address in enumerate(addresses):
        try:
            addresses[i] = formatAddress(address)
        except:
            continue
    used = []
    for add in addresses:
        if add not in used:
            used.append(add)
    addresses = used
    print("| Done.")

    ans = input("\n> Job ended. Do you want to save your contacts? yes/no  ")

    if ans == "yes":
        location = askdirectory(title="Chose you savings folder")
        chdir(location)
        with open("contacts.csv", "w") as file:
            for add in addresses:
                file.write(f"{add},\n")
        print("| Contacts saved.")

    ans = input("\n> Do you want to look further? yes/no  ")
    if ans == "yes":
        continue
    else:
        print("| Ending program. Closing opened mailboxes...")
        connectionClient.close()
        print("| Logging out...")
        connectionClient.logout()
        print("| Done. Program ended correctly. Thanks for using MailAddressesDownloader.")
        break

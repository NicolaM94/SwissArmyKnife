from getpass import getpass, getuser
import imaplib
from email import message_from_bytes
from pathlib import Path
from tkinter.filedialog import askdirectory
from os import chdir
import re

'''Mail Addresses downloader - Nicola Moro - 2022
Simply download addresses from your mail service and store them'''


TEMP_HOST = "imap.gmail.com"
TEMP_PORT = 993
TEMP_USERMAIL = "nicola.moro2312@gmail.com"
TEMPORARY_PASSWORD = "uymhxtywkdjflwhu"

def formatAddress (string) :

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
| Inserisci le informazioni richieste per scaricare tutti i tuoi contatti dall'indirizzo di posta.\n'''
)

hostUser = input("> Inserisci il server imap: ")
portUser = input("> Inserisci la porta imap: ")

connectionClient = None


print("\n| Connessione in corso...")
try:
    connectionClient = imaplib.IMAP4_SSL(host=TEMP_HOST,port=int(TEMP_PORT))
    print(f"| Connessione con {connectionClient.host} sulla porta {connectionClient.port} stabilita!")
except Exception as e:
    print(f"| Connessione con {connectionClient.host} sulla porta {connectionClient.port}: ",e.args)
    

print("\n| Inserisci le credenziali di accesso")
userMail = input("> Username: ")
userPassword = getpass(prompt="> Password: ")

print("| Login in corso...")
try:
    connectionClient.login(TEMP_USERMAIL,TEMPORARY_PASSWORD)
    print("| Login riuscito!")
except Exception as e:
    print("| Login fallito:",e.args)

while True:

    print("\n| Ecco le cartelle disponibili sul tuo account.")
    resp_code, directories = connectionClient.list()
    

    for n in directories:
        print(f"| ~ {n.decode()}")


    choice = input("\n> Indica da quale cartella pescare i contatti: ")


    typ, messages = connectionClient.select(f'"{choice}"')
    resp, mails = connectionClient.search(None,"ALL")
    mailsId = mails[0].decode().split()
    addresses = []

    print("| Mi metto all'opera. Ecco i contatti che arrivano...\n")
    for n in mailsId:
        resp, maildata = connectionClient.fetch(n,"(RFC822)")
        message = message_from_bytes(maildata[0][1])
        sender = message.get("From")
        receiver = message.get("To")
        if sender not in addresses:
            addresses.append(sender)
        if receiver not in addresses:
            addresses.append(receiver)
        print(f"From {sender} to {receiver}")


    print("\n|Solo un momento ora: provo a normalizzare gli indirizzi...")
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
    print("| Fatto!")


    ans = input("\n> Finito! Vuoi salvare i contatti? si/no  ")

    if ans == "si":
        location = askdirectory(title="Scegli dove salvare i contatti")
        chdir(location)
        with open("Contatti.csv","w") as file:
            for add in addresses:
                file.write(f"{add},\n")
        print("|Contatti salvati!")

    ans = input("\n> Vuoi cercare altri contatti? si/no  ")
    if ans == "si":
        continue
    else:
        print("| Ottimo, chiudo la connessione le caselle aperte...")
        connectionClient.close()
        print("| Effettuo il logout...")
        connectionClient.logout()
        print("| Tutto fatto! Alla prossima!")
        break



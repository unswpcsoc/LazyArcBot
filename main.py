#required packages
import discord
import datetime
import os
from dotenv import load_dotenv
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

#all secret keys and google drive fileId are store in the '.env' file
load_dotenv()
discord_key = os.getenv('Discord_BOT-TOKEN')
team_drive_id_UNSWPCsoc = os.getenv('ShareDriveFileID-UNSW_PCsoc')
parent_folder_id_21T3 = os.getenv('SubFolderFileID-21T3')
default_form_file_id = os.getenv('ARCFormFileID-Google_Forms')

#sets the varible to a text
default_form_title = "Arc Online Event Attendance List"
#sets the varible to whatever date is today
folder_naming_scheme = datetime.date.today().strftime("%Y-%m-%d")

#Google drive API authenication stuffs (pyDrive) with creditials stored in 'client_secrets.json' file
client = discord.Client()
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

#Discord API stuffs
@client.event
async def on_ready():
    #tell us when we have a successful login. example: "______#1234"
    print('We have logged in as {0.user}'.format(client))

@client.event
async def on_message(message):
    #makes sure the bot doesn't endlessly respond to itself
    if message.author == client.user:
        return

    #checks if new message starts with '~", if so the bot will run the code below
    if message.content.startswith("~"):
        #sets variable to whatever message is recieved
        InputTitle = message.content
        #sets the variable to original recived message but without the "~" character
        EventName = InputTitle.upper().replace("~","")

        #creates a folder for the new event, appropriate naming scheme of YYY-MM-DD followed my the event name, saves the folder in the right location
        NewEventFolder = drive.CreateFile({
            'title': folder_naming_scheme + " " + EventName,
            'mimeType': "application/vnd.google-apps.folder",
            'parents': [{
                'kind': 'drive#fileLink',
                'teamDriveId': team_drive_id_UNSWPCsoc,
                'id': parent_folder_id_21T3,
            }]
        })
        NewEventFolder.Upload()

        print("New Folder:  " + 'title: %s, id: %s' % (NewEventFolder['title'], NewEventFolder['id']))
        #sets the variable to whatever the newly created folder's ID is
        NewEventFolder_id = NewEventFolder['id']

        #copies the template google forms file and saves it to the newly created event folder
        NewEventForm = drive.auth.service.files().copy(
            fileId=default_form_file_id,
            body={
                "parents": [{
                    "kind": "drive#fileLink",
                    "id": NewEventFolder_id
                }],
                'title': default_form_title,
            }
        ).execute()

        print("New Form:    " + 'title: %s, id: %s' % (NewEventForm['title'],NewEventForm['id']))
        #sets the varible to whatever the newly copied form's ID is
        NewEventForm_id = NewEventForm['id']
        #sets the variable to a text that is formated correctly to work as a URL
        NewEventForm_link = "https://docs.google.com/forms/d/" + NewEventForm_id
        print(NewEventFolder['title'])
        print(NewEventForm_link)

        await message.channel.send("here is the new event ARC form link")
        #replys back to the orginal message with the form's URL
        await message.channel.send(NewEventForm_link)
client.run(discord_key)


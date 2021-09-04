#required packages
import discord
import datetime
import os
from dotenv import load_dotenv
from discord.ext import commands
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

#all secret keys and google drive fileId are store in the '.env' file
load_dotenv()
discord_key = os.getenv('Discord_BOT-TOKEN')
discord_role_id_EXEC = os.getenv('DiscordRoleID_EXEC')
discord_role_id_MOD = os.getenv('DiscordRoleID_MOD')
discord_channel_id = os.getenv('DiscordChannelID_execs')
team_drive_id_UNSWPCsoc = os.getenv('ShareDriveFileID-UNSW_PCsoc')
parent_folder_id_21T3 = os.getenv('SubFolderFileID-21T3')
default_form_file_id = os.getenv('ARCFormFileID-Google_Forms')

#sets the varible to a text
default_form_title = "Arc Online Event Attendance List"
#sets the varible to whatever date is today
folder_naming_scheme = datetime.date.today().strftime("%m-%d-%Y")
#sets the text string from "env" file to a intager numbers so that EXEC and MOD roles will have access
discord_role_id_EXEC = int(discord_role_id_EXEC)
discord_role_id_MOD = int(discord_role_id_MOD)
discord_channel_id = int(discord_channel_id)

#Discord API stuffs
client = discord.Client()
@client.event
async def on_ready():
    #tell us when we have a successful login. example: "______#1234"
    print('We have logged in to Discord as {0.user}'.format(client))

@client.event
@commands.guild_only()
async def on_message(message):
    #makes sure the bot doesn't endlessly respond to itself
    if message.author == client.user:
        return
    
    #ensures that bot or google drive isn't broken when and title begins with a " " or ([space]).
    if message.content.startswith("~arc  "):
        return
    
    #checks if new message starts with '~", if so the bot will run the code below
    if message.content.startswith("~arc ") and message.channel.id == discord_channel_id and any(i.id == discord_role_id_MOD or discord_role_id_EXEC for i in message.author.roles):
        #sets variable to whatever message is recieved
        InputTitle = message.content
        #sets the variable to original recived message but without the "~" character
        EventName = InputTitle.replace("~arc ","").upper()
        #sets the variabel to whomever sent the initial meesage
        AuthorisedExec = message.author

        #Google drive API authenication stuffs (pyDrive) with creditials stored in 'client_secrets.json' file. also the google authenication is put here so that the keys don't expire... i think. so now it authenticates everytime
        gauth = GoogleAuth()
        gauth.LocalWebserverAuth()
        drive = GoogleDrive(gauth)
        
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
        print(AuthorisedExec)
        print(NewEventForm_link)

        #replys back to the orginal message with the form's URL
        await message.reply(NewEventForm_link)
client.run(discord_key)


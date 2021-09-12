#required packages
import discord
import datetime
import time
import os
import pandas
from dotenv import load_dotenv
from discord.ext import commands
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from selenium import webdriver

#all secret keys and google drive fileId are store in the '.env' file
load_dotenv()
discord_key = os.getenv('Discord_BOT-TOKEN')
discord_role_id_EXEC = os.getenv('DiscordRoleID_EXEC')
discord_role_id_MOD = os.getenv('DiscordRoleID_MOD')
discord_channel_id = os.getenv('DiscordChannelID_execs')
team_drive_id_UNSWPCsoc = os.getenv('ShareDriveFileID-UNSW_PCsoc')
parent_folder_id_21T3 = os.getenv('SubFolderFileID-21T3')
default_form_file_id = os.getenv('ARCFormFileID-Google_Forms')
ClubName = os.getenv('ClubName_xPath')
FirstName = os.getenv('ApplicantFirstName')
LastName = os.getenv('ApplicantLastName')
EmailAddress = os.getenv('ApplicantEmail')
ContactNumber = os.getenv('ApplicantContactNumber')

#sets the varible to a text
default_form_title = "Arc Online Event Attendance List"
#sets the text string from "env" file to a intager numbers so that EXEC and MOD roles will have access
discord_role_id_EXEC = int(discord_role_id_EXEC)
discord_role_id_MOD = int(discord_role_id_MOD)
discord_channel_id = int(discord_channel_id)
                               
#Discord API stuffs
client = discord.Client()
@client.event
@commands.guild_only()
async def on_ready():
    #tell us when we have a successful login. example: "______#1234"
    print('We have logged in to Discord as {0.user}'.format(client))
    
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
        
        #sets the varible to whatever date is today
        folder_naming_scheme = datetime.date.today().strftime("%m-%d-%Y")

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

#this part is all for the Spac Grants function
def Sparc ():
   #get google drive api to download the google sheets as .xlsx and .pdf
    
    places = []
    # open file and read the content in a list
    with open("NameList.txt", 'r') as filehandle:
        places = [current_place.rstrip() for current_place in filehandle.readlines()]
        print(places)
    matched_indexes = []
    i = 0
    length = len(places)
    while i < length:
        if ClubName == places[i]:
            matched_indexes.append(i + 1)
        i += 1
    print(f'{ClubName} is present in list at indexes {matched_indexes}')
    ClubName_xPath = f'//*[@id="field91696630"]/option{matched_indexes}'
    print(ClubName_xPath)
    
    local_arc_files_path = r'C:\Users\ray\Desktop\arc_forms'
    testPATH = f'{local_arc_files_path} + {folder_naming_scheme} + .xlsx'
    form_sheet = pandas.read_excel(r'{testPATH}')
    
    # PATHa = "C:\\tools\selenium\MicrosoftWebDriver.exe"
    PATH = "C:\Program Files (x86)\Selenium\chromedriver.exe"

    browser = webdriver.Chrome(PATH)
    browser.get('https://arclimited.formstack.com/forms/clubfunding_onlineactivities')
    # web.close()

    print('loading web page in 5 sec')
    time.sleep(5)
    print('page should have loaded')

    AgreeConditions = "Yes"
    conditions = browser.find_element_by_xpath('//*[@id="field91696651_1"]')
    conditions.click()

    ClubName = "Computer Enthusiasts Society"
    club = browser.find_element_by_xpath(ClubName_xPath)
    club.click()

    fname = browser.find_element_by_xpath('//*[@id="field91696631-first"]')
    fname.send_keys(FirstName)

    lname = browser.find_element_by_xpath('//*[@id="field91696631-last"]')
    lname.send_keys(LastName)

    email = browser.find_element_by_xpath('//*[@id="field91696632"]')
    email.send_keys(ApplicantEmail)

    mobile = browser.find_element_by_xpath('//*[@id="field91696633"]')
    mobile.send_keys(ContactNumber)

    ConductedConjunction = "No"
    collaboration = browser.find_element_by_xpath('//*[@id="field91696634_2"]')
    collaboration.click()

    ActvityName = "?"
    activity = browser.find_element_by_xpath('//*[@id="field91698228"]')
    activity.send_keys(ActvityName)

    # def StartDate()
    smonth = web.find_element_by_xpath('//*[@id="field91698232M"]')
    smonth.send_keys(StartDate)
    sday = web.find_element_by_xpath('//*[@id="field91698232D"]')
    sday.click()
    syear = web.find_element_by_xpath('//*[@id="field91698232Y"]')
    shour = web.find_element_by_xpath('//*[@id="field91698232H"]')
    sminute = web.find_element_by_xpath('//*[@id="field91698232I"]')
    sampm = web.find_element_by_xpath('//*[@id="field91698232A"]')

    EndDate = "?"
    emonth = web.find_element_by_xpath('//*[@id="field91698235M"]')
    eday = web.find_element_by_xpath('//*[@id="field91698235D"]')
    eyear = web.find_element_by_xpath('//*[@id="field91698235Y"]')
    ehour = web.find_element_by_xpath('//*[@id="field91698235H"]')
    eminute = web.find_element_by_xpath('//*[@id="field91698235I"]')
    eampm = web.find_element_by_xpath('//*[@id="field91698235I"]')

    DescriptionActivity = "?"
    description = web.find_element_by_xpath('//*[@id="field91698239"]')
    description.send_keys(DescriptionActivity)

    PeopleAttended = "?"
    people = web.find_element_by_xpath('//*[@id="field91696639"]')
    people.send_keys(PeopleAttended)

    EvidenceAttendance = "?file?"
    evidence = web.find_element_by_xpath('//*[@id="field91698387UploadButton"]')
    evidence.send_keys(os.getcwd(local_arc_files_path + "\\" + folder_naming_scheme + ".pdf"))

    ActivityPhoto = "?file?"
    photo = web.find_element_by_xpath('//*[@id="field91698320UploadButton"]')
    photo.send_keys(os.getcwd(local_arc_files_path + "\\" + folder_naming_scheme + ".jpg"))
    
    IncomeExpenditure = "No"
    money = web.find_element_by_xpath('//*[@id="field91698326_2"]')
    money.send_click()

    NotesComments = "Powered my Lazy Arc Bot AI | @tld8102"
    notes = web.find_element_by_xpath('//*[@id="field91696649"]')
    notes.send_keys(NotesComments)

    SubmitForm = "Yes"
    submit = web.find_element_by_xpath('//*[@id="fsSubmitButton3856301"]')
    # submit.click()

    Confirmation = web.find_element_by_css_selector('.[class]')
    print(Confirmation.text)
    if ((Confirmation.text) == "?"):
        print("Successfully Submitted")
    else:
        print("Script Failed to Submit Form")
    
